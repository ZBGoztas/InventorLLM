import pythoncom
from win32com.client import dynamic

unk = pythoncom.GetActiveObject("Inventor.Application")
app = dynamic.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))
doc = app.ActiveDocument
cd = doc.ComponentDefinition
p = cd.Parameters.Item("d3")

print(f"Önce : {p.Expression} | Value = {p.Value} cm | Kütle = {cd.MassProperties.Mass:.3f} kg")

tr = app.TransactionManager.StartTransaction(doc, "LLM test: d3")
try:
    p.Expression = "170 mm"
    doc.Update()
    print(f"Sonra: {p.Expression} | Value = {p.Value} cm | Kütle = {cd.MassProperties.Mass:.3f} kg")
    input("Inventor'da değişikliğe bak, geri almak için Enter'a bas...")
finally:
    tr.Abort()

doc.Update()
print(f"Geri : {cd.Parameters.Item('d3').Expression}")