import datetime, json, os, pathlib, subprocess, sys, time, traceback, uuid
import pythoncom, win32api, win32con, win32event, win32gui, win32process, win32security
import win32com.client
from openpyxl import load_workbook

ROOT = pathlib.Path(__file__).resolve().parent
BOOK = ROOT / "dummy.xlsm"
RESULT = ROOT / "vba-result.txt"
LOG = ROOT / "worker.jsonl"
RUN = ROOT / "run.json"

def stamp():
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")

def event(name, **kw):
    with LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(dict(time=stamp(), event=name, **kw), ensure_ascii=False) + "\n")

def excel_processes():
    cmd = "Get-CimInstance Win32_Process -Filter \"Name = {}EXCEL.EXE{}\" | Select-Object ProcessId,ParentProcessId,CreationDate,ExecutablePath,CommandLine | ConvertTo-Json -Compress".format(chr(39), chr(39))
    r = subprocess.run(["powershell.exe", "-NoProfile", "-Command", cmd], capture_output=True, text=True, timeout=25, check=True)
    x = json.loads(r.stdout) if r.stdout.strip() else []
    return x if isinstance(x, list) else [x]

def owner(handle):
    token = win32security.OpenProcessToken(handle, win32con.TOKEN_QUERY)
    sid = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
    name, domain, _ = win32security.LookupAccountSid(None, sid)
    return domain + "\\" + name

def windows(pid):
    out = []
    def callback(hwnd, _):
        if win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.IsWindowVisible(hwnd):
            out.append({"hwnd":hwnd,"title":win32gui.GetWindowText(hwnd),"class":win32gui.GetClassName(hwnd)})
    win32gui.EnumWindows(callback, None)
    return out

def worker(run_id, old_pids):
    pythoncom.CoInitialize()
    app = wb = ph = None
    owned = False
    try:
        event("worker_start", pid=os.getpid(), run_id=run_id)
        app = win32com.client.DispatchEx("Excel.Application")
        hwnd = int(app.Hwnd)
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        if pid in old_pids:
            event("unsafe_shared_instance", pid=pid)
            app = None
            raise RuntimeError("DispatchEx returned a pre-existing Excel PID; no changes made")
        ph = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
        meta = [x for x in excel_processes() if x.get("ProcessId") == pid]
        if len(meta) != 1 or pathlib.Path(meta[0].get("ExecutablePath") or "").name.lower() != "excel.exe":
            event("unsafe_process_identity", pid=pid, metadata=meta)
            app = None
            raise RuntimeError("Excel PID identity not established; not closing it")
        owned = True
        event("excel_started", pid=pid, hwnd=hwnd, owner=owner(ph), process=meta[0], creation_time=str(win32process.GetProcessTimes(ph)["CreationTime"]), version=str(app.Version), automation_security=int(app.AutomationSecurity), workbooks=int(app.Workbooks.Count))
        if app.Workbooks.Count != 0:
            event("unexpected_startup_workbooks", action="left instance untouched")
            owned = False
            app = None
            raise RuntimeError("Unexpected startup workbooks; instance left open")
        if not app.EnableEvents:
            raise RuntimeError("Disabled Excel events; stopped")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Add()
        wb.Worksheets(1).Range("A2").Value = "R8"
        wb.Worksheets(1).Range("A3").Value = run_id
        code = (ROOT / "ThisWorkbook.bas").read_text(encoding="utf-8")
        wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString(code)
        wb.SaveAs(str(BOOK), FileFormat=52)
        if pathlib.Path(wb.FullName).resolve() != BOOK.resolve():
            raise RuntimeError("Workbook path mismatch")
        event("initial_saved", path=str(BOOK), a1_before=wb.Worksheets(1).Range("A1").Value, input=wb.Worksheets(1).Range("A2").Value)
        wb.Close(SaveChanges=False)
        wb = None
        event("opening_for_workbook_open", path=str(BOOK))
        wb = app.Workbooks.Open(str(BOOK), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False)
        if pathlib.Path(wb.FullName).resolve() != BOOK.resolve():
            raise RuntimeError("Opened workbook path mismatch")
        event("open_returned", result_exists=RESULT.exists())
        wb.Close(SaveChanges=False)
        wb = None
        app.Quit()
        app = None
        exited = win32event.WaitForSingleObject(ph, 15000) == win32event.WAIT_OBJECT_0
        event("excel_exit", pid=pid, exited=exited)
        return 0 if exited else 2
    except Exception:
        event("worker_exception", error=traceback.format_exc())
        return 1
    finally:
        if owned and app is not None:
            try:
                if wb is not None: wb.Close(SaveChanges=False)
                app.Quit()
                exited = win32event.WaitForSingleObject(ph, 15000) == win32event.WAIT_OBJECT_0
                event("cleanup_excel_exit", exited=exited)
            except Exception:
                event("cleanup_error", error=traceback.format_exc())
        pythoncom.CoUninitialize()

def main():
    if len(sys.argv) > 1 and sys.argv[1] == "worker":
        return worker(sys.argv[2], json.loads(sys.argv[3]))
    if any(p.exists() for p in (RUN, BOOK, RESULT, LOG)):
        raise RuntimeError("One attempt only per directory; existing artifacts must not be reused")
    run_id = uuid.uuid4().hex
    before = excel_processes()
    data = {"task":"BUD-161", "run_id":run_id, "started":stamp(), "attempt":1, "input":"R8", "approval_request":"not applicable to dummy job", "excel_before":before, "mode":"Workbook_Open", "timeout_seconds":90}
    with RUN.open("x", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    with (ROOT / "worker-output.log").open("x", encoding="utf-8") as out:
        proc = subprocess.Popen([sys.executable, str(pathlib.Path(__file__).resolve()), "worker", run_id, json.dumps([x["ProcessId"] for x in before])], stdout=out, stderr=subprocess.STDOUT)
        deadline = time.monotonic() + 90
        observed = []
        while proc.poll() is None and time.monotonic() < deadline:
            if LOG.exists():
                for line in LOG.read_text(encoding="utf-8").splitlines():
                    try: item = json.loads(line)
                    except json.JSONDecodeError: continue
                    if item.get("event") == "excel_started":
                        for w in windows(item["pid"]):
                            if w not in observed: observed.append(w)
            time.sleep(0.25)
        timed_out = proc.poll() is None
        if timed_out:
            proc.terminate()
            try: proc.wait(timeout=10)
            except subprocess.TimeoutExpired: pass
        code = proc.poll()
    entries = [json.loads(x) for x in LOG.read_text(encoding="utf-8").splitlines()] if LOG.exists() else []
    result = RESULT.read_text(encoding="cp932").strip() if RESULT.exists() else None
    actual = None
    read_error = None
    if BOOK.exists():
        try:
            book = load_workbook(BOOK, read_only=True, data_only=False)
            actual = {"a1":book.worksheets[0]["A1"].value, "a2":book.worksheets[0]["A2"].value, "a3":book.worksheets[0]["A3"].value}
            book.close()
        except Exception:
            read_error = traceback.format_exc()
    after = excel_processes()
    exit_event = any(x.get("event") == "excel_exit" and x.get("exited") for x in entries)
    started = [x for x in entries if x.get("event") == "excel_started"]
    pid = started[0]["pid"] if started else None
    survived = any(x.get("ProcessId") == pid for x in after) if pid else None
    ok = not timed_out and code == 0 and result is not None and result.startswith(run_id + "|ok|R8|") and actual == {"a1":"R8・処理完了", "a2":"R8", "a3":run_id} and exit_event and survived is False and not any(w["class"] == "#32770" for w in observed)
    data.update(ended=stamp(), worker_exit=code, timed_out=timed_out, observed_excel_windows=observed, approval_observation="No approval dialogs observed in periodic visible-window sampling" if not observed else "Visible Excel windows observed; inspect titles", vba_result=result, workbook_values=actual, workbook_read_error=read_error, excel_pid=pid, excel_process_handle_exit_confirmed=exit_event, excel_pid_present_after=survived, excel_after=after, verdict="PASS" if ok else ("FAIL" if result and ("|error|" in result or (actual and actual.get("a1") != "R8・処理完了")) else "INCONCLUSIVE"))
    RUN.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0 if ok else 1

if __name__ == "__main__":
    raise SystemExit(main())
