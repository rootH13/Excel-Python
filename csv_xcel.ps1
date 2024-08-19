$pythonProgramPath = cmd /c "where python" '2>&1'
$pythonScriptPath = "C:\your_path_to_\CSV2XCEL.py"
$pythonOutput = & $pythonProgramPath $pythonScriptPath $arg
$pythonOutput