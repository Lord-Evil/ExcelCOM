import clr
import System
clr.AddReference("C:\\Program Files (x86)\\Microsoft Office\\Office16\\DCF\\Microsoft.Office.Interop.Excel.dll")
import Microsoft.Office.Interop.Excel as Excel
clr.AddReference("Microsoft.VisualBasic")
import Microsoft.VisualBasic

from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal

def TypeName(obj):
    if(isinstance(obj, comobj)):
        obj=obj.obj
    return Microsoft.VisualBasic.Information.TypeName(obj)

COM=System.__ComObject
class comobj(object):
    def __init__(self, obj):
        self.obj = obj
        self.typ = obj.GetType()
        self.name = TypeName(obj)
        if(type(obj)==COM):
            self._class=None
            try:
                self._class=getattr(Excel, self.name)(obj)
            except:
                pass
        else:
            self._class=obj
    def __setattr__(self, name, value):
        if(name not in ["obj","typ", "name","_class","n"]):
            self.typ.InvokeMember(name, BindingFlags.Instance | BindingFlags.SetProperty, None, self.obj, [value])
        else: super(comobj, self).__setattr__(name, value)
    def __getattr__(self, name):
        method=None
        if(self._class is not None):
            if(hasattr(self._class, name)):
                ret=getattr(self._class, name)
                if(not isinstance(ret, COM)):
                    if(str(type(ret))=="<class 'CLR.MethodBinding'>"):
                        method=ret
                    else:
                        return ret
                else:
                    return comobj(ret)
        def newm(*argsv):
            noArgs=False
            if not argsv:
               argsv=(None,)
               noArgs=True
            else:
                argsv=list(argsv)
                argsv = [x.obj if type(x)==comobj else x for x in argsv]
                argsv=tuple(argsv)
            try:
                if(method):
                    if(noArgs):
                        result=method()
                    else:
                        result=method(*argsv)
                else:
                    result = self.typ.InvokeMember(name,BindingFlags.InvokeMethod | BindingFlags.GetProperty, None, self.obj, *argsv)
            except:
                print("Call was: "+name+" "+str(argsv))
                raise
            if(type(result)==COM):
                return comobj(result)
            else:
                return result
        return newm
    def __getitem__(self, idx):
        if(self._class is not None):
            if(hasattr(self._class, "get_Item")):
                if(type(idx)!=tuple):
                    idx=(idx,)
                ret=getattr(self._class, "get_Item")(*idx)
                if(type(ret)!=COM):
                    return ret
                else:
                    return comobj(ret)
            elif(hasattr(self._class, "Item")):
                if(type(idx)!=tuple):
                    idx=(idx,)
                ret=getattr(self._class, "Item")(*idx)
                if(type(ret)!=COM):
                    return ret
                else:
                    return comobj(ret)
        raise Error("Indexes not found")
    def __setitem__(self, idx, value):
        pass #todo: implemet this
    def __iter__(self):
        self.n = self.GetEnumerator()
        return self

    def __next__(self):
        if(self.n.MoveNext()):
            ret=self.n.Current
            if(type(ret)!=COM):
                return ret
            else:
                return comobj(ret)
        else:
            raise StopIteration
    def __str__(self):
        if(self._class is not None):
            return "COMObject("+self.name+")"
        else:
            return "Unknown COMObject"

def InspectObject(obj):
    inspectedObject={
        "Name":obj.name,
        "props":[]
    }
    for x in dir(obj._class):
        if not x.startswith("_"):
            try:
                ret = getattr(obj._class,x)
                if(str(type(ret))=="<class 'CLR.MethodBinding'>"):
                    #inspectedObject["methods"].append({x:ret})
                    pass
                else:
                    inspectedObject["props"].append({x:ret})
            except:
                inspectedObject["props"].append({x: "__EXCEPTION__"})
    return inspectedObject

appExcel=None
def GetExcel():
    global appExcel
    try:
        appExcel = comobj(Marshal.GetActiveObject("Excel.Application"))
    except:
        appExcel = comobj(Excel.ApplicationClass())
    if(not appExcel.Visible): appExcel.Visible=True
    if(not appExcel.Workbooks.Count): appExcel.Workbooks.Add()
    return appExcel

def QuitExcel():
    global appExcel
    appExcel.Quit();
    Marshal.ReleaseComObject(appExcel.obj)
    appExcel = None


if(__name__=="__main__"):
    app = GetExcel()
    workbook = app.ActiveWorkbook
    sheet=workbook.ActiveSheet
    sheet.Rows[1].Cells[1].Value="Bingo!!!"

'''
#usage example:

import ExcelCOM
app = ExcelCOM.GetExcel()
workbook = app.ActiveWorkbook
sheet=workbook.ActiveSheet
#general formatting are simple:
sheet.Rows[3].Cells[4].Value2=10.11
sheet.Rows[3].Cells[4].Value2*=4
print(sheet.Rows[3].Cells[4].Value2)
print(sheet.Rows[3].Cells[4].Value())

#DateTimes formatting is more interesting:
sheet.Rows[2].Cells[4].NumberFormat='[$-F800]dddd, mmmm dd, yyyy'
sheet.Rows[2].Cells[4].Value2=39521
print(sheet.Rows[2].Cells[4].Value2)
print(sheet.Rows[2].Cells[4].Value())

#Getting values from selected Range:
values = [v for v in app.ActiveWindow.RangeSelection.Value2]
print(values)
'''