#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

Global $excelFile = @ScriptDir & "\dieuchuyenve.xlsm"

;============================
; TẠO FORM
;============================

$gui = GUICreate("NHẬP THÔNG TIN ĐỐI TƯỢNG", 800, 650)

GUICtrlCreateLabel("Họ và tên",20,20)
$b7 = GUICtrlCreateInput("",150,20,200)

GUICtrlCreateLabel("Tên gọi khác",20,50)
$c8 = GUICtrlCreateInput("",150,50,200)

GUICtrlCreateLabel("Ngày tháng năm sinh",20,80)
$c9 = GUICtrlCreateInput("",150,80,200)

GUICtrlCreateLabel("Đăng ký HKTT",20,110)
$b11 = GUICtrlCreateInput("",150,110,200)

GUICtrlCreateLabel("Dân tộc",20,140)
$c12 = GUICtrlCreateInput("",150,140,200)

GUICtrlCreateLabel("Hành vi phạm tội",20,170)
$c13 = GUICtrlCreateInput("",150,170,200)

GUICtrlCreateLabel("Bắt ngày",20,200)
$c14 = GUICtrlCreateInput("",150,200,200)

GUICtrlCreateLabel("Vào TTG ngày",20,230)
$c15 = GUICtrlCreateInput("",150,230,200)

GUICtrlCreateLabel("Án phạt",20,260)
$c16 = GUICtrlCreateInput("",150,260,200)

GUICtrlCreateLabel("Giới tính",20,290)
$e9 = GUICtrlCreateInput("",150,290,200)

;============================
; THÔNG TIN BẢN ÁN
;============================

GUICtrlCreateLabel("Lệnh số ",350,20)
$b21 = GUICtrlCreateInput("",470,20,150)

GUICtrlCreateLabel("Ngày ra ",350,65)
$c21 = GUICtrlCreateInput("",470,65,150)

GUICtrlCreateLabel("CQ ra lệnh",350,110)
$d21 = GUICtrlCreateInput("",470,110,150)

GUICtrlCreateLabel("Cán bộ giao",350,155)
$e21 = GUICtrlCreateInput("",470,155,150)

GUICtrlCreateLabel("Cấp bậc",350,190)
$f21 = GUICtrlCreateInput("",470,190,150)

GUICtrlCreateLabel("Đơn vị",350,230)
$g21 = GUICtrlCreateInput("",470,230,150)




;============================
; BUTTON
;============================

$btnSave = GUICtrlCreateButton("LƯU",320,380,170,60)
$btnPrintGNDC = GUICtrlCreateButton("In Giao nhận (Điều chuyển)",180,480,170,40)
$btnPrintBTGGDC = GUICtrlCreateButton("In BTGG (Điều chuyển)",180,580,170,40)
$btnPrintGNTX = GUICtrlCreateButton("In Giao nhận (Trích xuất)",450,480,170,40)
$btnPrintBTGGTX = GUICtrlCreateButton("In BTGG (Trích xuất)",450,580,170,40)




GUISetState()

;============================
; LOOP
;============================

While 1
    $msg = GUIGetMsg()

    Switch $msg
        Case $GUI_EVENT_CLOSE
            Exit

        Case $btnSave
            SaveExcel()

        Case $btnPrintGNDC
            SaveExcel()
            PrintSheet("GN-DC")

        Case $btnPrintBTGGDC
            SaveExcel()
            PrintSheet("BTGG-DC")

        Case $btnPrintGNTX
            SaveExcel()
            PrintSheet("GN-TX")

        Case $btnPrintBTGGTX
            SaveExcel()
            PrintSheet("BTGG-TX")

        		
    EndSwitch
WEnd

;============================
; HÀM LƯU
;============================

Func SaveExcel()

    Local $oExcel = ObjCreate("Excel.Application")
    $oExcel.Visible = False

    Local $oWorkbook = $oExcel.Workbooks.Open($excelFile)
    Local $oSheet = $oWorkbook.Sheets("TTCHUNG")

    $oSheet.Range("B7").Value = GUICtrlRead($b7)
    $oSheet.Range("C8").Value = GUICtrlRead($c8)
    $oSheet.Range("C9").Value = GUICtrlRead($c9)
    $oSheet.Range("B11").Value = GUICtrlRead($b11)
    $oSheet.Range("C12").Value = GUICtrlRead($c12)
    $oSheet.Range("C13").Value = GUICtrlRead($c13)
    $oSheet.Range("C14").Value = GUICtrlRead($c14)
    $oSheet.Range("C15").Value = GUICtrlRead($c15)
    $oSheet.Range("C16").Value = GUICtrlRead($c16)
    $oSheet.Range("E9").Value = GUICtrlRead($e9)

    $oSheet.Range("B21").Value = GUICtrlRead($b21)
    $oSheet.Range("C21").Value = GUICtrlRead($c21)
    $oSheet.Range("D21").Value = GUICtrlRead($d21)
    $oSheet.Range("E21").Value = GUICtrlRead($e21)
    $oSheet.Range("F21").Value = GUICtrlRead($f21)
    $oSheet.Range("g21").Value = GUICtrlRead($g21)

    $oWorkbook.Save()
    $oWorkbook.Close()
    $oExcel.Quit()

EndFunc

;============================
; HÀM IN (ĐÃ SỬA IN 2 MẶT)
;============================

Func PrintSheet($sheetName)

    Local $oExcel = ObjCreate("Excel.Application")
    $oExcel.Visible = False

    Local $oWorkbook = $oExcel.Workbooks.Open($excelFile)
    Local $oSheet = $oWorkbook.Sheets($sheetName)

    ; ===== BẬT IN 2 MẶT =====
    ; 2 = lật cạnh dài (chuẩn)
    $oSheet.PageSetup.Duplex = 2

    ; ===== IN =====
    $oSheet.PrintOut

    $oWorkbook.Close(False)
    $oExcel.Quit()

EndFunc