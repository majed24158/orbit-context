"""Offline safety checks for the owned process and untrusted response boundary."""
import sys,time,tempfile,json
from pathlib import Path
from unittest.mock import patch
import win32api,win32process,win32job,win32con
from codex_bridge import _hidden_process,run_codex
from document_conversion import read_text
from file_tools import ConversionCancelled

checks=[]
root=Path(__file__).parent/'qa'; root.mkdir(exist_ok=True)
with tempfile.TemporaryDirectory(prefix='bridge-check-',dir=root) as temporary:
    folder=Path(temporary); stdin=folder/'input'; stdin.write_bytes(b'')
    log=folder/'process.log'; started=time.monotonic()
    try:
        _hidden_process([sys.executable,'-c','import os,time; print(os.getpid(),flush=True); time.sleep(60)'],
            stdin,log,folder,lambda:time.monotonic()-started>.5,10)
    except ConversionCancelled: pass
    else: raise AssertionError('process cancellation did not propagate')
    assert time.monotonic()-started<8
    pid=int(log.read_text().strip())
    try: handle=win32api.OpenProcess(0x1000,False,pid)
    except Exception: pass
    else:
        try: assert win32process.GetExitCodeProcess(handle)!=win32con.STILL_ACTIVE
        finally: handle.Close()
    checks.append('cancellation terminates the owned hidden process Job promptly')
    with (patch('win32job.AssignProcessToJobObject',side_effect=RuntimeError('assignment rejected')),
          patch('win32process.TerminateProcess',wraps=win32process.TerminateProcess) as terminate):
        try: _hidden_process([sys.executable,'-c','import time; time.sleep(60)'],stdin,log,folder,None,5)
        except RuntimeError: pass
        else: raise AssertionError('failed Job assignment was ignored')
        assert terminate.call_count==1
    checks.append('failed Job assignment terminates only its suspended process before execution')
    def fake_process(arguments,stdin_path,log_path,*unused):
        answer=Path(arguments[arguments.index('-o')+1])
        answer.write_text('{"content":"safe"}',encoding='utf-8')
        Path(log_path).write_text(json.dumps({'item':{'type':'command_execution'}}),encoding='utf-8')
        return 0
    with patch('codex_bridge._hidden_process',side_effect=fake_process):
        try: run_codex('synthetic',{},[],folder)
        except RuntimeError as error: assert 'أداة' in str(error)
        else: raise AssertionError('unexpected tool execution was accepted')
    checks.append('unexpected CLI tool event rejects the result')
    def malformed(arguments,stdin_path,log_path,*unused):
        Path(arguments[arguments.index('-o')+1]).write_text('{"content":"a","content":"b"}',encoding='utf-8')
        Path(log_path).write_text('',encoding='utf-8'); return 0
    with patch('codex_bridge._hidden_process',side_effect=malformed):
        try: run_codex('synthetic',{},[],folder)
        except ValueError: pass
        else: raise AssertionError('duplicate JSON keys accepted')
    source=folder/'large.txt'
    with source.open('wb') as handle: handle.truncate(10_000_000)
    try: read_text(source)
    except ValueError: pass
    else: raise AssertionError('oversized TXT accepted')
    checks.append('duplicate JSON keys and oversized TXT are rejected')
report={'ok':True,'count':len(checks),'checks':checks,'network_used':False}
(root/'bridge-check.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report,indent=2))
