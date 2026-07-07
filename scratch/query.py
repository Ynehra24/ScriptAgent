import subprocess

def test_script(script):
    import tempfile, os
    with tempfile.NamedTemporaryFile(mode='w', suffix='.applescript', delete=False) as f:
        f.write(script)
        tmp = f.name
    try:
        res = subprocess.run(["osascript", tmp], capture_output=True, text=True)
        return res.stdout.strip(), res.stderr.strip(), res.returncode
    finally:
        os.unlink(tmp)

script = '''
tell application "Microsoft Word"
    try
        select text object of paragraph 1 of document 1
        set style of selection to style "Heading 1" of document 1
        return "Success"
    on error e
        return "Runtime Error: " & e
    end try
end tell
'''
out, err, code = test_script(script)
print("Code:", code)
print("Out:", out)
print("Err:", err)
