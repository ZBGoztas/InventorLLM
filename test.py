import pythoncom
from win32com.client import dynamic

unk = pythoncom.GetActiveObject("Inventor.Application")
app = dynamic.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))
print("Inventor:", app.SoftwareVersion.DisplayVersion)

doc = app.ActiveDocument
print("Aktif doküman:", doc.DisplayName)

for p in doc.ComponentDefinition.Parameters:
    print(f"  {p.Name:20} {p.Expression}")