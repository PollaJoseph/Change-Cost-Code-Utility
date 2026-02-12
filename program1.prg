*-- 1. Setup Environment --*
PUBLIC oForm
oForm = CREATEOBJECT("PayITStyleForm")
oForm.Show
READ EVENTS

DEFINE CLASS PayITStyleForm AS Form
    Caption = "Material Cost Code Utility"
    WindowState = 0
    Width = 900
    Height = 720
    MinWidth = 700
    MinHeight = 650
    FontName = "Segoe UI"
    BorderStyle = 3
    BackColor = RGB(248, 249, 250) 
    Themes = .F.
    
    nConnHandle = 0

    PROCEDURE Load
        LOCAL cConnString, nResult, cChartSQL
        
        cConnString = "Driver={SQL Server};" + ;
                      "Server=SERVER_NAME;" + ;
                      "Database=DATABASE_NAME;" + ;
                      "Uid=USERNAME;" + ;
                      "Pwd=USER_PASSWORD;"
        
        THIS.nConnHandle = SQLSTRINGCONNECT(cConnString)
        
        IF THIS.nConnHandle < 0
            MESSAGEBOX("Connection Failed to SQL Server", 16, "SQL Error")
            RETURN .F.
        ENDIF

        *-- Initial SQL Data Fetching
        SQLEXEC(THIS.nConnHandle, "SELECT cProjectCode, iProjectId FROM dbo.tblProject ORDER BY cProjectCode", "curProjects")
        
        cChartSQL = "SELECT (accode + curcod) as full_code, ldesc FROM [DATABASE_NAME].[dbo].[chart] WHERE lactive = 1 ORDER BY accode"
        SQLEXEC(THIS.nConnHandle, cChartSQL, "curChart")

        *-- Initialize placeholders
        CREATE CURSOR curPOs (ordno C(20))
        CREATE CURSOR curItems (poitem C(10), acode C(20))
        CREATE CURSOR curSRVs (mrvno C(20))
        CREATE CURSOR curSRVItems (mrvitem C(10), acode C(20))
    ENDPROC

    PROCEDURE Destroy
        IF THIS.nConnHandle > 0
            SQLDISCONNECT(THIS.nConnHandle)
        ENDIF
        CLEAR EVENTS
    ENDPROC
    
    PROCEDURE Resize
        THIS.shpHeader.Width = THIS.Width
        IF TYPE("THIS.cntMain") = "O"
            WITH THIS.cntMain
                .Left = (ThisForm.Width - .Width) / 2
                .Top  = (ThisForm.Height - .Height) / 2 + 30 
                IF .Top < 80 
                    .Top = 80
                ENDIF
            ENDWITH
        ENDIF
    ENDPROC

    *-- UI HEADER COMPONENTS --*
    ADD OBJECT shpHeader AS Shape WITH Top = 0, Left = 0, Height = 60, Width = 900, ;
        BackColor = RGB(40, 55, 70), BorderStyle = 0, Anchor = 10
        
        *-- LOGO SETTINGS --------------------------------------------------------    
    ADD OBJECT imgLogo AS Image WITH ;
        Left = 20, ;
        Height = 60, ;
        Width = 60, ;
        Stretch = 1, ;           
        BackStyle = 0, ;     
        Picture = "I:\TALISMAN10\Update Cost Code Utility\CCC.png"
        
    ADD OBJECT lblAppTitle AS Label WITH Caption = "Material Cost Code Utility", ;
        Top = 15, Left = 70, FontName = "Segoe UI", FontSize = 14, FontBold = .T., ;
        ForeColor = RGB(255, 255, 255), BackStyle = 0, AutoSize = .T.

    ADD OBJECT cntMain AS Container WITH Width = 600, Height = 620, BackStyle = 0, BorderWidth = 0

    PROCEDURE Init
        WITH THIS.cntMain
            *-- Row 1: Sub Project
            .AddObject("lblProj", "Label")
            .lblProj.Caption = "Project"
            .lblProj.Visible = .T.
            
            .AddObject("cmbProject", "ComboBox")
            WITH .cmbProject
                .Top = 20
                .Width = 600
                .Style = 2
                .Visible = .T.
                .RowSourceType = 3
                .RowSource = "SELECT cProjectCode, iProjectId FROM curProjects INTO CURSOR cboProj NOFILTER"
            ENDWITH

            *-- Row 2: PO Details
            .AddObject("lblPO", "Label")
            WITH .lblPO
                .Caption = "PO Number "
                .Top = 60
                .Visible = .T.
            ENDWITH
            
            .AddObject("lblItem", "Label")
            WITH .lblItem
                .Caption = "PO Item "
                .Top = 60
                .Left = 310
                .Visible = .T.
            ENDWITH
            
            .AddObject("cmbPO", "ComboBox")
            WITH .cmbPO
                .Top = 80
                .Width = 290
                .Style = 2
                .Enabled = .F.
                .Visible = .T.
            ENDWITH
            
            .AddObject("cmbItem", "ComboBox")
            WITH .cmbItem
                .Top = 80
                .Left = 310
                .Width = 290
                .Style = 2
                .Enabled = .F.
                .Visible = .T.
            ENDWITH

            *-- Row 3: SRV Details
            .AddObject("lblSRV", "Label")
            WITH .lblSRV
                .Caption = "SRV Number "
                .Top = 120
                .Visible = .T.
            ENDWITH
            
            .AddObject("lblSRVItem", "Label")
            WITH .lblSRVItem
                .Caption = "SRV Item "
                .Top = 120
                .Left = 310
                .Visible = .T.
            ENDWITH
            
            .AddObject("cmbSRV", "ComboBox")
            WITH .cmbSRV
                .Top = 140
                .Width = 290
                .Style = 2
                .Enabled = .F.
                .Visible = .T.
            ENDWITH
            
            .AddObject("cmbSRVItem", "ComboBox")
            WITH .cmbSRVItem
                .Top = 140
                .Left = 310
                .Width = 290
                .Style = 2
                .Enabled = .F.
                .Visible = .T.
            ENDWITH

            *-- Row 4: DUAL COST CODE DISPLAY
            .AddObject("lblPOCost", "Label")
            WITH .lblPOCost
                .Caption = "PO Cost Code "
                .Top = 185
                .Visible = .T.
            ENDWITH
            
            .AddObject("txtPOCost", "TextBox")
            WITH .txtPOCost
                .Top = 205
                .Width = 290
                .ReadOnly = .T.
                .BackColor = RGB(240, 240, 240)
                .Visible = .T.
                .FontBold = .T.
            ENDWITH

            .AddObject("lblSRVCost", "Label")
            WITH .lblSRVCost
                .Caption = "SRV Cost Code "
                .Top = 185
                .Left = 310
                .Visible = .T.
            ENDWITH
            
            .AddObject("txtSRVCost", "TextBox")
            WITH .txtSRVCost
                .Top = 205
                .Left = 310
                .Width = 290
                .ReadOnly = .T.
                .BackColor = RGB(240, 240, 240)
                .Visible = .T.
                .FontBold = .T.
            ENDWITH

            *-- Row 5: New Cost Code
            .AddObject("lblNew", "Label")
            WITH .lblNew
                .Caption = "New Cost Code "
                .Top = 250
                .Visible = .T.
            ENDWITH
            
            .AddObject("cmbNewCode", "ComboBox")
            WITH .cmbNewCode
                .Top = 270
                .Width = 600
                .Style = 2
                .Visible = .T.
                .RowSourceType = 3
                .RowSource = "SELECT full_code, ldesc FROM curChart INTO CURSOR cboNewCodeList"
            ENDWITH

            *-- Buttons
            .AddObject("cmdUpdate", "CommandButton")
            WITH .cmdUpdate
                .Caption = "UPDATE RECORDS"
                .Left = 150
                .Top = 340
                .Width = 300
                .Height = 45
                .BackColor = RGB(152, 251, 152)
                .Themes = .F.
                .Visible = .T.
            ENDWITH
            
            .AddObject("cmdExit", "CommandButton")
            WITH .cmdExit
                .Caption = "EXIT SYSTEM"
                .Left = 150
                .Top = 395
                .Width = 300
                .Height = 45
                .BackColor = RGB(255, 160, 122)
                .Themes = .F.
                .Visible = .T.
            ENDWITH
        ENDWITH
        
        *-- Bind Events
        BINDEVENT(THIS.cntMain.cmbProject, "InteractiveChange", THIS, "OnProjectChange")
        BINDEVENT(THIS.cntMain.cmbPO, "InteractiveChange", THIS, "OnPOChange")
        BINDEVENT(THIS.cntMain.cmbItem, "InteractiveChange", THIS, "OnItemChange")
        BINDEVENT(THIS.cntMain.cmbSRV, "InteractiveChange", THIS, "OnSRVChange")
        BINDEVENT(THIS.cntMain.cmbSRVItem, "InteractiveChange", THIS, "OnSRVItemChange")
        BINDEVENT(THIS.cntMain.cmdUpdate, "Click", THIS, "OnSaveClick")
        BINDEVENT(THIS.cntMain.cmdExit, "Click", THIS, "OnExitClick")
        
        THIS.Resize()
    ENDPROC

    *-- LOGIC METHODS --*
    
    PROCEDURE OnProjectChange
        LOCAL cProjCode, cSQL
        cProjCode = ALLTRIM(ThisForm.cntMain.cmbProject.Value)
        IF EMPTY(cProjCode)
            RETURN
        ENDIF
        
        cSQL = "SELECT DISTINCT ordno FROM dbo.pmspod WHERE cProjectCode = ?cProjCode AND superceded = 0 ORDER BY ordno"
        SQLEXEC(ThisForm.nConnHandle, cSQL, "curPOs")
        
        WITH ThisForm.cntMain
            .cmbPO.RowSourceType = 3
            .cmbPO.RowSource = "select ordno from curPOs into cursor cboPOsFinal"
            .cmbPO.Enabled = .T.
            .cmbPO.Requery()
            .cmbPO.ListIndex = 0
            
            .cmbItem.Enabled = .F.
            .txtPOCost.Value = ""
            .txtSRVCost.Value = ""
        ENDWITH
    ENDPROC

    PROCEDURE OnPOChange
        LOCAL cPONum, cSQL
        cPONum = ALLTRIM(cboPOsFinal.ordno)
        
        cSQL = "SELECT poitem, acode FROM dbo.pmspod WHERE ordno = ?cPONum AND superceded = 0 ORDER BY poitem"
        SQLEXEC(ThisForm.nConnHandle, cSQL, "curItems")
        
        WITH ThisForm.cntMain
            .cmbItem.RowSourceType = 3
            .cmbItem.RowSource = "select poitem, acode from curItems into cursor cboItemsFinal"
            .cmbItem.Enabled = .T.
            .cmbItem.Requery()
            .cmbItem.ListIndex = 0
            
            .txtPOCost.Value = ""
            .txtSRVCost.Value = ""
        ENDWITH
    ENDPROC
    
    PROCEDURE OnItemChange
        LOCAL cPONum, cPOItem, cSQL
        cPONum  = ALLTRIM(cboPOsFinal.ordno)
        cPOItem = ALLTRIM(cboItemsFinal.poitem)
        
        *-- Display PO Cost Code
        ThisForm.cntMain.txtPOCost.Value = cboItemsFinal.acode 
        
        cSQL = "SELECT DISTINCT mrvno FROM dbo.pmssrvd WHERE ordno = ?cPONum AND poitem = ?cPOItem AND revno = 0 ORDER BY mrvno"
        SQLEXEC(ThisForm.nConnHandle, cSQL, "curSRVs")
        
        WITH ThisForm.cntMain
            .cmbSRV.RowSourceType = 3
            .cmbSRV.RowSource = "select mrvno from curSRVs into cursor cboSRVsFinal"
            .cmbSRV.Enabled = .T.
            .cmbSRV.Requery()
            .cmbSRV.ListIndex = 0
            
            .txtSRVCost.Value = ""
        ENDWITH
    ENDPROC
    
    PROCEDURE OnSRVChange
        LOCAL cPONum, cPOItem, cMRVNo, cSQL
        cPONum  = ALLTRIM(cboPOsFinal.ordno)
        cPOItem = ALLTRIM(cboItemsFinal.poitem)
        cMRVNo  = ALLTRIM(cboSRVsFinal.mrvno)
        
        cSQL = "SELECT mrvitem, acode FROM dbo.pmssrvd WHERE mrvno = ?cMRVNo AND ordno = ?cPONum AND poitem = ?cPOItem AND revno = 0 ORDER BY mrvitem"
        SQLEXEC(ThisForm.nConnHandle, cSQL, "curSRVItems")
        
        WITH ThisForm.cntMain
            .cmbSRVItem.RowSourceType = 3
            .cmbSRVItem.RowSource = "select mrvitem, acode from curSRVItems into cursor cboSRVItemsFinal"
            .cmbSRVItem.Enabled = .T.
            .cmbSRVItem.Requery()
            .cmbSRVItem.ListIndex = 0
        ENDWITH
    ENDPROC

    PROCEDURE OnSRVItemChange
        *-- Display SRV Cost Code
        ThisForm.cntMain.txtSRVCost.Value = cboSRVItemsFinal.acode  
    ENDPROC

    PROCEDURE OnSaveClick
        LOCAL cNewCode, cOrdNo, cPOItem, cMRVNo, cMRVItem, cSQL, nResPO, nResSRV
        cNewCode = ALLTRIM(ThisForm.cntMain.cmbNewCode.DisplayValue)
        
        IF EMPTY(cNewCode)
            MESSAGEBOX("Please select a New Cost Code first.", 48, "Validation")
            RETURN
        ENDIF
        
        cOrdNo = ALLTRIM(cboPOsFinal.ordno)
        cPOItem = ALLTRIM(cboItemsFinal.poitem)
        
        SQLSETPROP(ThisForm.nConnHandle, 'Transactions', 2) 
        
        *-- Update PO
        cSQL = "UPDATE [DATABASE_NAME].[dbo].[pmspod] SET acode = ?cNewCode WHERE ordno = ?cOrdNo AND poitem = ?cPOItem AND superceded = 0"
        nResPO = SQLEXEC(ThisForm.nConnHandle, cSQL)

        *-- Update SRV if selected
        IF ThisForm.cntMain.cmbSRV.ListIndex > 0 AND ThisForm.cntMain.cmbSRVItem.ListIndex > 0
            cMRVNo = ALLTRIM(cboSRVsFinal.mrvno)
            cMRVItem = ALLTRIM(cboSRVItemsFinal.mrvitem)
            
            cSQL = "UPDATE [DATABASE_NAME].[dbo].[pmssrvd] SET acode = ?cNewCode WHERE mrvno = ?cMRVNo AND mrvitem = ?cMRVItem AND revno = 0"
            nResSRV = SQLEXEC(ThisForm.nConnHandle, cSQL)
        ELSE
            nResSRV = 1
        ENDIF

        IF nResPO > 0 AND nResSRV > 0
            SQLCOMMIT(ThisForm.nConnHandle)
            MESSAGEBOX("Update Successful!", 64, "Success")
            
            *-- Refresh Displays
            ThisForm.cntMain.txtPOCost.Value = cNewCode
            IF !EMPTY(cMRVNo) 
                ThisForm.cntMain.txtSRVCost.Value = cNewCode
            ENDIF
        ELSE
            SQLROLLBACK(ThisForm.nConnHandle)
            MESSAGEBOX("Error occurred during update.", 16, "Error")
        ENDIF
        SQLSETPROP(ThisForm.nConnHandle, 'Transactions', 1) 
    ENDPROC

    PROCEDURE OnExitClick
        IF MESSAGEBOX("Are you sure you want to exit?", 4 + 32, "Exit System") = 6 
            ThisForm.Release() 
        ENDIF
    ENDPROC
ENDDEFINE