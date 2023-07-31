# PowerShell script to execute two Python files sequentially for the current user

# Get the currently logged-in user's name
$currentUserName = $env:USERNAME

# Construct the path to the user-specific Python executable
$pythonExe = "C:\Users\$currentUserName\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\python.exe"

# Paths to your Python files
$pythonFile1 = "C:\Users\eyilmazdemir\Desktop\barcode\excel\excel_test.py"
$pythonFile2 = "C:\Users\eyilmazdemir\Desktop\barcode\barcodegui.pyw"

# Execute the first Python file
Start-Process -FilePath $pythonExe -ArgumentList $pythonFile1 -Wait -WindowStyle Hidden

# Execute the second Python file after the first one completes
Start-Process -FilePath $pythonExe -ArgumentList $pythonFile2 -Wait

Write-Host "ALLONS-Y!"
