import sys, os, winreg, glob, pythoncom
import win32com.client

VBA_MODULE = '''Option Explicit
Dim st As New Collection

Sub TogglePic(ByVal sname As String)
    Dim shp As Shape, sw As Double, sh As Double, ar As Double
    Dim saved As String, v As Variant
    For Each shp In ActivePresentation.SlideShowWindow.View.Slide.Shapes
        If shp.Name = sname Then
            sw = ActivePresentation.PageSetup.SlideWidth
            sh = ActivePresentation.PageSetup.SlideHeight
            On Error Resume Next
            saved = st(sname)
            On Error GoTo 0
            If saved = "" Then
                st.Add shp.Left & "|" & shp.Top & "|" & shp.Width & "|" & shp.Height, sname
                ar = shp.Width / shp.Height
                If ar > sw / sh Then
                    shp.Width = sw: shp.Height = sw / ar
                    shp.Top = (sh - shp.Height) / 2: shp.Left = 0
                Else
                    shp.Height = sh: shp.Width = sh * ar
                    shp.Left = (sw - shp.Width) / 2: shp.Top = 0
                End If
                shp.ZOrder msoBringToFront
            Else
                v = Split(saved, "|")
                shp.Left = CDbl(v(0)): shp.Top = CDbl(v(1))
                shp.Width = CDbl(v(2)): shp.Height = CDbl(v(3))
                st.Remove sname
            End If
            Exit For
        End If
    Next
End Sub

Sub RestoreAllOnEnd(Pres As Presentation)
    Dim s As Slide, shp As Shape, v As Variant
    For Each s In Pres.Slides
        For Each shp In s.Shapes
            If shp.Type = 13 Then
                On Error Resume Next
                v = Split(st(shp.Name), "|")
                On Error GoTo 0
                If IsArray(v) Then
                    shp.Left = CDbl(v(0)): shp.Top = CDbl(v(1))
                    shp.Width = CDbl(v(2)): shp.Height = CDbl(v(3))
                End If
            End If
        Next
    Next
End Sub
'''

VBA_EVENTCLASS = '''Public WithEvents PPTApp As Application

Private Sub PPTApp_SlideShowEnd(ByVal Pres As Presentation)
    RestoreAllOnEnd Pres
End Sub
'''


def reg_trust(setting):
    k = r'Software\Microsoft\Office\16.0\PowerPoint\Security'
    try:
        h = winreg.CreateKey(winreg.HKEY_CURRENT_USER, k)
        winreg.SetValueEx(h, 'VBAWarnings', 0, winreg.REG_DWORD, 1 if setting else 0)
        winreg.SetValueEx(h, 'AccessVBOM', 0, winreg.REG_DWORD, 1 if setting else 0)
        winreg.CloseKey(h)
    except Exception:
        pass


def set_module(cm, code):
    cl = cm.CountOfLines
    if cl > 0:
        cm.DeleteLines(1, cl)
    cm.AddFromString(code)


def process_file(path):
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        print(f'  [SKIP] Not found: {path}')
        return

    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.ppt', '.pptx', '.pptm'):
        print(f'  [SKIP] Unsupported: {ext}')
        return

    is_pptx = ext == '.pptx'
    pres = None
    try:
        app = win32com.client.Dispatch('PowerPoint.Application')
        pres = app.Presentations.Open(path)

        vba_code = VBA_MODULE
        idx = 0
        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for shi in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes(shi)
                if shp.Type == 13:
                    shp.Name = f'SS_{idx}'
                    shp.ActionSettings(1).Run = f'TF_{idx}'
                    vba_code += f'\nSub TF_{idx}()\n    TogglePic "SS_{idx}"\nEnd Sub\n'
                    idx += 1

        vba_code += '''
Dim evt As EventClass

Sub AutoOpen()
    Set evt = New EventClass
    Set evt.PPTApp = Application
End Sub
'''

        vbproj = pres.VBProject
        try:
            mod = vbproj.VBComponents('Module1')
        except Exception:
            mod = vbproj.VBComponents.Add(1)
        set_module(mod.CodeModule, vba_code)

        try:
            ec = vbproj.VBComponents('EventClass')
        except Exception:
            ec = vbproj.VBComponents.Add(2)
            ec.Name = 'EventClass'
        set_module(ec.CodeModule, VBA_EVENTCLASS)

        if is_pptx:
            new_path = os.path.splitext(path)[0] + '.pptm'
            pres.SaveAs(new_path, 24)  # ppSaveAsOpenXMLMacroEnabled
            print(f'  [OK] {idx} pictures -> {os.path.basename(new_path)}')
        else:
            pres.Save()
            print(f'  [OK] {idx} pictures: {os.path.basename(path)}')

        pres.Close()
    except Exception as e:
        print(f'  [ERR] {os.path.basename(path)}: {e}')
        if pres:
            try:
                pres.Close()
            except Exception:
                pass


def main():
    reg_trust(True)

    if len(sys.argv) >= 2:
        # Single file mode
        pythoncom.CoInitialize()
        try:
            process_file(sys.argv[1])
        finally:
            pythoncom.CoUninitialize()
    else:
        # Batch mode: scan script dir
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        files = []
        for ext in ('*.ppt', '*.pptx', '*.pptm'):
            files.extend(glob.glob(os.path.join(script_dir, ext)))
        files = sorted(set(files))

        if not files:
            print(f'No .ppt/.pptx/.pptm found in:\n  {script_dir}')
            exit(1)

        print(f'Found {len(files)} file(s):')
        pythoncom.CoInitialize()
        try:
            for f in files:
                process_file(f)
        finally:
            pythoncom.CoUninitialize()

    reg_trust(False)


if __name__ == '__main__':
    main()
