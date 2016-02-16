VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm SANPROJECT 
   BackColor       =   &H8000000C&
   Caption         =   "SAN'S PROJECT"
   ClientHeight    =   9165
   ClientLeft      =   165
   ClientTop       =   1635
   ClientWidth     =   10575
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu log 
      Caption         =   "&Log"
      Begin VB.Menu log_logOn 
         Caption         =   "Log O&n "
         Shortcut        =   ^N
      End
      Begin VB.Menu log_logOff 
         Caption         =   "Log O&ff"
         Shortcut        =   ^O
      End
      Begin VB.Menu blank1 
         Caption         =   "-"
      End
      Begin VB.Menu log_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu administration 
      Caption         =   "&Administration"
      Begin VB.Menu admProEntry 
         Caption         =   "Product Entry"
      End
      Begin VB.Menu admProUpd 
         Caption         =   "Product Update "
      End
      Begin VB.Menu blank19 
         Caption         =   "-"
      End
      Begin VB.Menu adm_createUser 
         Caption         =   "Create &User"
         Shortcut        =   ^U
      End
      Begin VB.Menu adm_modifyUser 
         Caption         =   "&ModifY User"
         Shortcut        =   ^M
      End
      Begin VB.Menu adm_ChangePassword 
         Caption         =   "C&hange Password"
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
      End
      Begin VB.Menu adm_backUp 
         Caption         =   "&Back UP"
         Shortcut        =   ^B
      End
      Begin VB.Menu adm_recovery 
         Caption         =   "&Recovery"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu purchase 
      Caption         =   "&Purchase"
      Begin VB.Menu pur_ven 
         Caption         =   "Vendor"
         Begin VB.Menu pur_ven_Entry 
            Caption         =   "&Entry"
         End
         Begin VB.Menu pur_ven_Update 
            Caption         =   "&Update"
         End
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu pur_purEntry 
         Caption         =   "Purchase &Entry"
      End
      Begin VB.Menu pur_purReturn 
         Caption         =   "Purchase &Return"
      End
   End
   Begin VB.Menu sale 
      Caption         =   "&Sale "
      Begin VB.Menu sale_cust 
         Caption         =   "&Customer"
         Begin VB.Menu sale_cust_Entry 
            Caption         =   "&Entry"
         End
         Begin VB.Menu sale_cust_Update 
            Caption         =   "&Update"
         End
      End
      Begin VB.Menu blank4 
         Caption         =   "-"
      End
      Begin VB.Menu sale_party 
         Caption         =   "&Party"
         Begin VB.Menu sale_party_Entry 
            Caption         =   "&Entry"
         End
         Begin VB.Menu sale_party_Update 
            Caption         =   "&Update"
         End
      End
      Begin VB.Menu blank5 
         Caption         =   "-"
      End
      Begin VB.Menu sale_saleEntry 
         Caption         =   "Sale &Entry"
      End
      Begin VB.Menu sale_saleReturn 
         Caption         =   "Sale &Return"
      End
   End
   Begin VB.Menu repla 
      Caption         =   "R&eplacement"
      Begin VB.Menu repl_on_spt 
         Caption         =   "On Spot Replacement"
      End
      Begin VB.Menu blank113 
         Caption         =   "-"
      End
      Begin VB.Menu repl_repl_from 
         Caption         =   "Replacement From"
      End
      Begin VB.Menu blank14 
         Caption         =   "-"
      End
      Begin VB.Menu repl_repl_to_prin 
         Caption         =   "Replace To Principal"
      End
      Begin VB.Menu blank15 
         Caption         =   "-"
      End
      Begin VB.Menu repl_defc_pro 
         Caption         =   "Defective Product  Entry"
      End
   End
   Begin VB.Menu empSuport 
      Caption         =   "&Employee Support"
      Begin VB.Menu emp_joinEmp 
         Caption         =   "&Join Employee"
         Begin VB.Menu emp_join_entry 
            Caption         =   "&Entry"
         End
         Begin VB.Menu emp_join_update 
            Caption         =   "&Update"
         End
      End
      Begin VB.Menu blank6 
         Caption         =   "-"
      End
      Begin VB.Menu emp_relEmp 
         Caption         =   "&Relieving Employee"
      End
      Begin VB.Menu blank7 
         Caption         =   "-"
      End
      Begin VB.Menu emp_sal 
         Caption         =   "Employee &Salary"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu masterDetail 
      Caption         =   "&Master Detail"
      Begin VB.Menu mas_stk_detail 
         Caption         =   "Stock Detail"
      End
      Begin VB.Menu bank9 
         Caption         =   "-"
      End
      Begin VB.Menu mas_pur_detail 
         Caption         =   "Purchase Detail "
      End
      Begin VB.Menu blank61 
         Caption         =   "-"
      End
      Begin VB.Menu mas_sale_detail 
         Caption         =   "Sale Detail "
      End
      Begin VB.Menu blank62 
         Caption         =   "-"
      End
      Begin VB.Menu mas_onSpotReplacementDetail 
         Caption         =   "On Spot Replacement Detail"
      End
      Begin VB.Menu blank63 
         Caption         =   "-"
      End
      Begin VB.Menu mas_replaceToPrnplDetail 
         Caption         =   "Replace To Principal Detail"
      End
      Begin VB.Menu blank64 
         Caption         =   "-"
      End
      Begin VB.Menu mas_replaceFromCust 
         Caption         =   "Replace from Customer"
      End
      Begin VB.Menu blank65 
         Caption         =   "-"
      End
      Begin VB.Menu mas_currEmpDetail 
         Caption         =   "Current Employee Detail"
      End
      Begin VB.Menu blank66 
         Caption         =   "-"
      End
      Begin VB.Menu mas_relieveEmpDetail 
         Caption         =   "Relieved Employee Detail"
      End
      Begin VB.Menu blank70 
         Caption         =   "-"
      End
      Begin VB.Menu mas_userHistory 
         Caption         =   "User History"
      End
   End
   Begin VB.Menu report 
      Caption         =   "&Report"
      Begin VB.Menu rep_pro 
         Caption         =   "Product Report"
         Begin VB.Menu rep_pro_al 
            Caption         =   "All Product"
         End
         Begin VB.Menu blank16 
            Caption         =   "-"
         End
         Begin VB.Menu rep_pro_name 
            Caption         =   "By Product Name"
         End
         Begin VB.Menu blank17 
            Caption         =   "-"
         End
         Begin VB.Menu rep_pro_cat 
            Caption         =   "By Category"
         End
         Begin VB.Menu blank18 
            Caption         =   "-"
         End
         Begin VB.Menu rep_pro_date 
            Caption         =   "By date"
         End
      End
      Begin VB.Menu blank20 
         Caption         =   "-"
      End
      Begin VB.Menu rep_ven 
         Caption         =   "Vendor Report"
         Begin VB.Menu rep_ven_all 
            Caption         =   "All Vendor"
         End
         Begin VB.Menu blank21 
            Caption         =   "-"
         End
         Begin VB.Menu rep_ven_add 
            Caption         =   "By Address"
         End
         Begin VB.Menu blank22 
            Caption         =   "-"
         End
         Begin VB.Menu rep_ven_delr_of 
            Caption         =   "By Deler of"
         End
      End
      Begin VB.Menu blank23 
         Caption         =   "-"
      End
      Begin VB.Menu rep_party 
         Caption         =   "Party Report"
         Begin VB.Menu rep_party_all 
            Caption         =   "All Party"
         End
         Begin VB.Menu blank24 
            Caption         =   "-"
         End
         Begin VB.Menu rep_party_nam 
            Caption         =   "By Name"
         End
      End
      Begin VB.Menu blank25 
         Caption         =   "-"
      End
      Begin VB.Menu rep_cust 
         Caption         =   "Customer Report"
         Begin VB.Menu rep_cust_all 
            Caption         =   "All Customer"
         End
         Begin VB.Menu blank26 
            Caption         =   "-"
         End
         Begin VB.Menu rep_cust_nam 
            Caption         =   "By Name "
         End
         Begin VB.Menu blank27 
            Caption         =   "-"
         End
         Begin VB.Menu rep_cust_add 
            Caption         =   "By Address"
         End
      End
      Begin VB.Menu blank28 
         Caption         =   "-"
      End
      Begin VB.Menu rep_pur 
         Caption         =   "Purchase Report "
         Begin VB.Menu rep_pur_sln 
            Caption         =   "By Purchase Sln"
         End
         Begin VB.Menu blank34 
            Caption         =   "-"
         End
         Begin VB.Menu rep_pur_ven 
            Caption         =   "By Vendor"
         End
         Begin VB.Menu blank29 
            Caption         =   "-"
         End
         Begin VB.Menu rep_pur_date 
            Caption         =   "By Date "
         End
      End
      Begin VB.Menu blank31 
         Caption         =   "-"
      End
      Begin VB.Menu rep_sale 
         Caption         =   "Sale Report"
         Begin VB.Menu rep_sale_cust 
            Caption         =   "By Customer"
         End
         Begin VB.Menu blank32 
            Caption         =   "-"
         End
         Begin VB.Menu rep_sale_party 
            Caption         =   "By Party"
         End
      End
      Begin VB.Menu blank36 
         Caption         =   "-"
      End
      Begin VB.Menu rep_stk 
         Caption         =   "Stock Report"
         Begin VB.Menu rep_stk_all 
            Caption         =   "All Product"
         End
         Begin VB.Menu blank44 
            Caption         =   "-"
         End
         Begin VB.Menu rep_stk_nam 
            Caption         =   "By Product Name"
         End
         Begin VB.Menu blank45 
            Caption         =   "-"
         End
         Begin VB.Menu rep_stk_cate 
            Caption         =   "By Category"
         End
         Begin VB.Menu blank46 
            Caption         =   "-"
         End
         Begin VB.Menu rep_stk_LTDate 
            Caption         =   "By Last Transaction Date"
         End
      End
      Begin VB.Menu blank43 
         Caption         =   "-"
      End
      Begin VB.Menu rep_rpl 
         Caption         =   "Replacement  Report"
         Begin VB.Menu rep_rpl_cust 
            Caption         =   "Replace from Customer"
         End
         Begin VB.Menu blank37 
            Caption         =   "-"
         End
         Begin VB.Menu rep_rpl_prnl 
            Caption         =   "Replace To Principal"
         End
         Begin VB.Menu blank38 
            Caption         =   "-"
         End
         Begin VB.Menu rep_rpl_on_spot 
            Caption         =   "On Spot Replacement "
         End
      End
      Begin VB.Menu blank39 
         Caption         =   "-"
      End
      Begin VB.Menu rep_defPro 
         Caption         =   "Defective Product Report"
         Begin VB.Menu rep_defPro_All 
            Caption         =   "All Defective Prodcut"
         End
         Begin VB.Menu blank40 
            Caption         =   "-"
         End
         Begin VB.Menu rep_defPro_Nam 
            Caption         =   "By Defective Product Name "
         End
         Begin VB.Menu blank41 
            Caption         =   "-"
         End
         Begin VB.Menu rep_defPro_cat 
            Caption         =   "By Defective Product Category"
         End
         Begin VB.Menu blank42 
            Caption         =   "-"
         End
         Begin VB.Menu rep_defPro_Date 
            Caption         =   "By Defective Product Entry Date"
         End
      End
      Begin VB.Menu blank50 
         Caption         =   "-"
      End
      Begin VB.Menu rep_emp 
         Caption         =   "Employee Report"
         Begin VB.Menu rep_emp_cEmp 
            Caption         =   "Current Employee "
            Begin VB.Menu rep_emp_cEmp_all 
               Caption         =   "All Current Employee"
            End
            Begin VB.Menu blank51 
               Caption         =   "-"
            End
            Begin VB.Menu rep_emp_cEmp_nam 
               Caption         =   "BY Name "
            End
            Begin VB.Menu blank52 
               Caption         =   "-"
            End
            Begin VB.Menu rep_emp_cEmp_doj 
               Caption         =   "By Date of Joining"
            End
         End
         Begin VB.Menu blank53 
            Caption         =   "-"
         End
         Begin VB.Menu rep_emp_rEmp 
            Caption         =   "Relieved Employee"
            Begin VB.Menu rep_emp_eEmp_All 
               Caption         =   "All Relieved Employee "
            End
            Begin VB.Menu blank54 
               Caption         =   "-"
            End
            Begin VB.Menu rep_emp_rEmp_nam 
               Caption         =   "By Name"
            End
            Begin VB.Menu blank55 
               Caption         =   "-"
            End
            Begin VB.Menu rep_emp_rEmp_doj 
               Caption         =   "By Relieved Employee Date of Joining"
            End
            Begin VB.Menu blank56 
               Caption         =   "-"
            End
            Begin VB.Menu rep_emp_rEmp_dor 
               Caption         =   "By Relieved Employee Date of Relieved"
            End
         End
         Begin VB.Menu blank57 
            Caption         =   "-"
         End
         Begin VB.Menu rep_emp_sal 
            Caption         =   "Employee Salary "
            Begin VB.Menu rep_emp_sal_all 
               Caption         =   "All Employee Salary"
            End
            Begin VB.Menu blank58 
               Caption         =   "-"
            End
            Begin VB.Menu rep_emp_sal_nam 
               Caption         =   "By Employee Name"
            End
            Begin VB.Menu blank59 
               Caption         =   "-"
            End
            Begin VB.Menu rep_emp_sal_month 
               Caption         =   "By Month of Payment"
            End
         End
      End
   End
   Begin VB.Menu fav 
      Caption         =   "&Favorites"
      Begin VB.Menu fav_calculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu fav_calender 
         Caption         =   "Calender"
      End
      Begin VB.Menu blank68 
         Caption         =   "-"
      End
   End
   Begin VB.Menu window 
      Caption         =   "&Window"
      Begin VB.Menu win_cascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu win_tileHorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu win_tileVerically 
         Caption         =   "Tile Verically"
      End
      Begin VB.Menu win_ArrangeIcons 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu help_about 
         Caption         =   "About San's Project "
      End
   End
   Begin VB.Menu pop_menu 
      Caption         =   "po"
      Visible         =   0   'False
      Begin VB.Menu pop_pur 
         Caption         =   "Purchase Entry"
      End
      Begin VB.Menu blank100 
         Caption         =   "-"
      End
      Begin VB.Menu pop_sale 
         Caption         =   "Sale Entry"
      End
      Begin VB.Menu blank101 
         Caption         =   "-"
      End
      Begin VB.Menu pop_stock 
         Caption         =   "Stock Detail"
      End
      Begin VB.Menu blank102 
         Caption         =   "-"
      End
      Begin VB.Menu pop_bg 
         Caption         =   "Background"
      End
   End
End
Attribute VB_Name = "SANPROJECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************'
'**                     **'
'** SANPROJECT MDI FORM **'
'**                     **'
'*************************'

'VARIABLE DECLARATION

Option Explicit

Dim odbca As String
Dim capName As String
Public hisSession As OraSession
Public hisDatabase As OraDatabase
Dim insertSql   As String

Private Sub adm_ChangePassword_Click()
    
    CHANGE_PASSWORD.Show

End Sub

Private Sub adm_createUser_Click()
    
    CREATE_USER.Show

End Sub

Private Sub adm_modifyUser_Click()
    
    MODIFY_USER.Show

End Sub

Private Sub admProEntry_Click()
    
    PRODUCT_ENTRY.Show

End Sub

Private Sub admProUpd_Click()
    
    PRODUCT_UPDATE.Show

End Sub

Private Sub emp_join_entry_Click()
    
    EMP_JOINING_ENTRY.Show

End Sub

Private Sub emp_join_update_Click()
    
    CURRENT_EMP_UPDATE.Show

End Sub

Private Sub emp_relEmp_Click()
    
    RELIEVING_EMPLOYEE.Show

End Sub

Private Sub emp_sal_Click()
    
    EMP_SALARY_DETAIL.Show

End Sub

Private Sub fav_calculator_Click()
    
    CALCULATOR.Show

End Sub

Private Sub help_about_Click()
    
    SanAbout.Show

End Sub

Private Sub log_exit_Click()
    
    insertSql = " insert into hrishi.sanproject_history values('" & ModuleVarious.LogOnUser & "','" & ModuleVarious.workingDate & "','" & ModuleVarious.Stime & "','" & Time & "')"
    hisDatabase.ExecuteSQL (insertSql)
    End

End Sub

Private Sub log_logOff_Click()
    
    log_logOn.Enabled = True
    log_logOff.Enabled = False
    administration.Enabled = False
    sale.Enabled = False
    purchase.Enabled = False
    repla.Enabled = False
    empSuport.Enabled = False
    masterDetail.Enabled = False
    report.Enabled = False
    fav.Enabled = False
    window.Enabled = False
    help.Enabled = False

    insertSql = " insert into hrishi.sanproject_history values('" & ModuleVarious.LogOnUser & "','" & ModuleVarious.workingDate & "','" & ModuleVarious.Stime & "','" & Time & "')"
    hisDatabase.ExecuteSQL (insertSql)

End Sub

Private Sub log_logOn_Click()
  
  
    LOG_ON.Show
 
End Sub

Private Sub mas_currEmpDetail_Click()

    MASTER_CURRENT_EMP_DETAIL.Show

End Sub

Private Sub mas_onSpotReplacementDetail_Click()

    MASTER_ONSPOT_REPLACE_DETAIL.Show

End Sub

Private Sub mas_pur_detail_Click()

    MASTER_PURCHASE_DETAIL.Show

End Sub

Private Sub mas_relieveEmpDetail_Click()
    
    MASTER_RELIVED_EMP_DETAIL.Show

End Sub

Private Sub mas_replaceFromCust_Click()
    
    MASTER_REPLACE_FROM_CUST_DETAIL.Show

End Sub

Private Sub mas_replaceToPrnplDetail_Click()

    MASTER_REPLACE_TO_PRINCIPAL_DETAIL.Show

End Sub

Private Sub mas_sale_detail_Click()
    
    MASTER_SALE_DETAIL.Show
End Sub

Private Sub mas_stk_detail_Click()

    STOCK_DETAIL.Show

End Sub

Private Sub mas_userHistory_Click()
    
    SANPROJECT_HISTORY.Show

End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu pop_menu

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    insertSql = " insert into hrishi.sanproject_history values('" & ModuleVarious.LogOnUser & "','" & ModuleVarious.workingDate & "','" & ModuleVarious.Stime & "','" & Time & "')"
    hisDatabase.ExecuteSQL (insertSql)
    If hisSession.LastServerErr = 0 Then
        If hisDatabase.LastServerErr = 0 Then
            If Err.Number = 0 Then
            Else
                MsgBox "VB Error: " & Err.Number & Err.Description, vbCritical, "VB Error:"
                Exit Sub
            End If
        Else
            MsgBox "Database Error:" & hisDatabase.LastServerErr & hisDatabase.LastServerErrText, vbCritical, "Database Error:"
            hisDatabase.LastServerErrReset
            Exit Sub
        End If
    Else
        MsgBox "Session Error : " & hisSession.LastServerErr & hisSession.LastServerErrText, vbCritical, "Sesssion Error:"
        hisSession.LastServerErrReset
        Exit Sub
    End If

End Sub

Private Sub pop_bg_Click()

Dim pictureName As String
CommonDialog1.InitDir = "e:/wallpapers"
CommonDialog1.Filter = "San's Project Images|*.BMP;*.GIF;*.JPG;*.DIB|All Files|*.*"
CommonDialog1.Action = 1
pictureName = CommonDialog1.FileName
If pictureName = "" Then Exit Sub
SANPROJECT.Picture = LoadPicture(pictureName)

End Sub

Private Sub pop_pur_Click()
  PUR_ENTRY.Show
End Sub

Private Sub pop_sale_Click()
   SALE_ENTRY.Show
End Sub

Private Sub pop_stock_Click()
STOCK_DETAIL.Show
End Sub

Private Sub pur_purEntry_Click()
    
    PUR_ENTRY.Show

End Sub

Private Sub pur_ven_Entry_Click()

    VEN_ENTRY.Show

End Sub

Private Sub pur_ven_Update_Click()

    VEN_UPDATE.Show

End Sub

Private Sub rep_currEmpDetail_Click()
    
    MASTER_CURRENT_EMP_DETAIL.Show

End Sub

Private Sub rep_cust_add_Click()

    Dim custAdd As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    custAdd = InputBox("Enter the customer Address :", "Customer Address Paramerter :")
    If custAdd = "" Then
        MsgBox "Sorry ! Customer Address is Blank .", vbExclamation, "Empty:"
        Exit Sub
    Else
        custAdd = UCase(custAdd)
    End If
    REPROTENVIRONMENT.custAddQry custAdd
    CUST_ADD_DETAIL.Show

End Sub

Private Sub rep_cust_all_Click()

    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.custAllDetail
    CUST_ALL_DETAIL.Show
     
End Sub

Private Sub rep_cust_nam_Click()

    Dim custName As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    custName = InputBox("Enter the name of Customer:", " Customer Name Parameter:")
    If custName = "" Then
        MsgBox "Sorry ! Customer name is Blank .", vbExclamation, "Empty:"
        Exit Sub
    Else
        custName = UCase(custName)
    End If
     REPROTENVIRONMENT.custNameQry custName
     CUST_NAME_DETAIL.Show
 
End Sub

Private Sub rep_defPro_All_Click()
    
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.defProAllDetail
    DEFECTIVE_PRO_ALL_DETAIL.Show

End Sub

Private Sub rep_defPro_cat_Click()

    Dim category As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    category = InputBox("ENTER THE PRODUCT NAME :", "PRODUCT NAME PARAMETER:")
    If category = "" Then
        MsgBox "SORRY ! PRODUCT CATEGORY IS EMPTY .", vbExclamation, "EMPTY:"
        Unload Me
    Else
        category = UCase(category)
   
    End If
    REPROTENVIRONMENT.defProCateQry category
    DEFECTIVE_PRO_CATEGORY_DETAIL.Show

End Sub

Private Sub rep_defPro_Date_Click()
    
    capName = rep_defPro_Date.Caption
    DATE_PARAMETER.Show

End Sub

Private Sub rep_defPro_Nam_Click()
    
    Dim proName As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    proName = InputBox("ENTER THE PRODUCT NAME :", "PRODUCT NAME PARAMETER:")
    If proName = "" Then
         MsgBox "SORRY ! PRODUCT NAME IS EMPTY .", vbExclamation, "EMPTY:"
         Unload Me
    Else
        proName = UCase(proName)
    End If
    REPROTENVIRONMENT.defProNameQry proName
    DEFECTIVE_PRO_NAME_DETAIL.Show
    
End Sub

Private Sub rep_emp_cEmp_all_Click()
    
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.crtEmpAllDetail
    CRT_EMP_DETAIL.Show
 
End Sub

Private Sub rep_emp_cEmp_doj_Click()
    
    capName = rep_emp_cEmp_doj.Caption
    DATE_PARAMETER.Show

End Sub

Private Sub rep_emp_cEmp_nam_Click()
    
    Dim empName As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    empName = InputBox("Enter the name of Employee .", "Name Parameter:")
    If empName = "" Then
        MsgBox "Sorry ! Employee isn't  present.", vbExclamation, "Employee:"
        Exit Sub
    Else
        empName = UCase(empName)
    End If
    REPROTENVIRONMENT.crtEmpNameQry empName
    CRT_EMP_NAME_DETAIL.Show
 
End Sub

Private Sub rep_emp_eEmp_All_Click()

    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.rvdEmpDetail
    RELIEVED_EMP_DETAIL.Show
 
End Sub

Private Sub rep_emp_rEmp_doj_Click()
    
    capName = rep_emp_rEmp_doj.Caption
    DATE_PARAMETER.Show

End Sub

Private Sub rep_emp_rEmp_dor_Click()

    capName = rep_emp_rEmp_dor.Caption
    DATE_PARAMETER.Show

End Sub

Private Sub rep_emp_rEmp_nam_Click()
    
    Dim rEmpName As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    rEmpName = InputBox("Enter Relieved Employee .", "Name Parameter.")
    If rEmpName = "" Then
        MsgBox "Sorry ! Relieved Employee Name is Blank .", vbExclamation, "Empty:"
        Exit Sub
    Else
        rEmpName = UCase(rEmpName)
    End If
    REPROTENVIRONMENT.rvdEmpNameQry rEmpName
    RELIEVED_EMP_NAME_DETAIL.Show

End Sub

Private Sub rep_emp_sal_all_Click()
    
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.salEmpDetail
    SALARY_EMP_DETAIL.Show
 
End Sub

Private Sub rep_emp_sal_month_Click()

    capName = rep_emp_sal_month.Caption
    MONTH_PARAMETER.Show

End Sub

Private Sub rep_emp_sal_nam_Click()

    Dim empName As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    empName = InputBox("Enter the Employee Name :", "Employee Name Parameter:")
    If empName = "" Then
        MsgBox "Sorry ! Employee Name is Empty.", vbExclamation, "Empty:"
        Exit Sub
    Else
        empName = UCase(empName)
    End If
    REPROTENVIRONMENT.salEmpNameQry empName
    SALARY_EMP_NAME_DETAIL.Show
 
End Sub

Private Sub rep_party_all_Click()

    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.partyAllDetail
    PARTY_ALL_DETAIL.Show
 
End Sub

Private Sub rep_party_nam_Click()

    Dim partyName As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    partyName = InputBox("Enter the party name :", "Party Name Parameter.")
    If partyName = "" Then
        MsgBox "Sorry! Party name is Blank.", vbExclamation, "Empty:"
        Exit Sub
    Else
        partyName = UCase(partyName)
    End If
    REPROTENVIRONMENT.partyNameQry partyName
    PARTY_NAME_DETAIL.Show

End Sub

Private Sub rep_pro_al_Click()

    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.proAllDetail
    PRO_ALL_DETAIL.Show

End Sub

Private Sub rep_pro_cat_Click()

    Dim category As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    category = InputBox("Enter the category.", "Search Parameter")
    If category = "" Then
        MsgBox "Sorry ! Category is blank.", vbInformation + vbOKOnly
        Exit Sub
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.proCateQry category
    PRO_CATE_DETAIL.Show
 
End Sub

Private Sub rep_pro_date_Click()
    
    capName = rep_pro_date.Caption
    DATE_PARAMETER.Show

End Sub

Private Sub rep_pro_name_Click()

    Dim proName As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    proName = InputBox("Ente the product .", "Search Parameter:")
    If proName = "" Then
        MsgBox "Sorry ! Product name is blank .", vbInformation + vbOKOnly
        Exit Sub
    Else
        proName = UCase(proName)
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.proNameQry proName
    PRO_NAME_DETAIL.Show
 
End Sub


Private Sub rep_pur_sln_Click()
    
    capName = rep_pur_sln.Caption
    PUR_SLN_PARAMETER.Show

End Sub

Private Sub rep_rpl_cust_Click()

    capName = rep_rpl_cust.Caption
    RPL_CUST_SLN_PARAMETER.Show

End Sub

Private Sub rep_rpl_prnl_Click()

    capName = rep_rpl_prnl.Caption
    RPL_PRNPAL_SLN_PARAMETER.Show

End Sub

Private Sub rep_sale_cust_Click()
    
    capName = rep_sale_cust.Caption
    SAL_SLN_CUST_PARAMETER.Show

End Sub

Private Sub rep_sale_detail_Click()
    
    MASTER_SALE_DETAIL.Show

End Sub

Private Sub rep_sale_party_Click()

    capName = rep_sale_party.Caption
    SAL_SLN_PARTY_PARAMETER.Show

End Sub

Private Sub rep_stk_all_Click()
    
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.stkAllDetail
    STOCK_ALL_DETAIL.Show

End Sub

Private Sub rep_stk_cate_Click()
    
    Dim category As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    category = InputBox("Enter the Product Category:", "Category Parameter:")
    If category = "" Then
        MsgBox "Sorry ! Category  is empty .", vbInformation, "Empty:"
        Exit Sub
    Else
        category = UCase(category)
    End If
    REPROTENVIRONMENT.stkCateQry category
    STOCK_CATEGORY_DETAIL.Show
    
End Sub

Private Sub rep_stk_LTDate_Click()
    
    capName = rep_stk_LTDate.Caption
    DATE_PARAMETER.Show

End Sub

Private Sub rep_stk_nam_Click()
    
    Dim proName As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    proName = InputBox("Enter the Product Name:", "Product Name Parameter:")
    If proName = "" Then
        MsgBox "Sorry ! Product Name is empty .", vbInformation, "Empty:"
        Exit Sub
    Else
        proName = UCase(proName)
    End If
    REPROTENVIRONMENT.stkProNameQry proName
    STOCK_PRONAME_DETAIL.Show

End Sub

Private Sub rep_ven_add_Click()
    
    Dim venAdd As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    venAdd = InputBox("Enter the address", "Address Parameter:")
    If venAdd = "" Then
        MsgBox "Sorry ! address is empty .", vbInformation, "Empty:"
        Exit Sub
    Else
        venAdd = UCase(venAdd)
    End If
    REPROTENVIRONMENT.venAddQry venAdd
    VEN_ADDRESS_DETAIL.Show

End Sub

Private Sub rep_ven_all_Click()


    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    REPROTENVIRONMENT.venAllDetail
    VEN_ALL_DETAIL.Show

End Sub

Private Sub rep_ven_delr_of_Click()
    
    Dim delerOf As String
    odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
    If REPROTENVIRONMENT.Connection.State = adStateOpen Then
        REPROTENVIRONMENT.Connection.Close
    End If
    REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
    delerOf = InputBox("Enter the Deler of vendor :", "Deler Of Parameter")
    If delerOf = "" Then
        MsgBox "Sorry ! The deler of is empty.", vbExclamation, "Empty"
        Exit Sub
    Else
        delerOf = UCase(delerOf)
    End If
    REPROTENVIRONMENT.venDelerOfQry delerOf
    VEN_DELEROF_DETAIL.Show

End Sub

Private Sub repl_defc_pro_Click()
    
    DEF_PRO_ENTRY.Show

End Sub

Private Sub repl_on_spt_Click()
    
    ON_SPOT_RLMNT.Show

End Sub

Private Sub repl_repl_from_Click()
    
    REPLACE_FROM.Show

End Sub

Private Sub repl_repl_to_prin_Click()
    
    REPLACE_TO_PRINCIPAL.Show

End Sub

Private Sub sale_cust_Entry_Click()

    CUSTOMER_ENTRY.Show

End Sub

Private Sub sale_cust_Update_Click()
    
    CUSTOMER_UPDATE.Show

End Sub

Private Sub sale_party_Entry_Click()
    
    PARTY_ENTRY.Show

End Sub

Private Sub sale_saleEntry_Click()

    SALE_ENTRY.Show

End Sub

Public Sub formShow()
     
     If capName = "By date" Then
        odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
        If REPROTENVIRONMENT.Connection.State = adStateOpen Then
            REPROTENVIRONMENT.Connection.Close
        End If
        REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
        REPROTENVIRONMENT.proEntryDate ModuleVarious.sDate, ModuleVarious.eDate
        PRO_ENTRYDATE_DETAIL.Show
     End If
 
     If capName = "By Purchase Sln" Then
         odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
        If REPROTENVIRONMENT.Connection.State = adStateOpen Then
            REPROTENVIRONMENT.Connection.Close
        End If
        REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
        REPROTENVIRONMENT.purAllDetail ModuleVarious.purSln
        PUR_ALL_DETAIL.Show
     End If
 
    If capName = "By Customer" Then
         odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
         If REPROTENVIRONMENT.Connection.State = adStateOpen Then
                REPROTENVIRONMENT.Connection.Close
         End If
         REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
         REPROTENVIRONMENT.saleCustDetail ModuleVarious.salSlnCust
         SALE_CUSTOMER_DETAIL.Show
    End If
 
    If capName = "By Party" Then
         odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
        If REPROTENVIRONMENT.Connection.State = adStateOpen Then
            REPROTENVIRONMENT.Connection.Close
        End If
        REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
        REPROTENVIRONMENT.salePartyDetail ModuleVarious.salSlnParty
        SALE_PARTY_DETAIL.Show
    End If
 
    If capName = "Replace from Customer" Then
        odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
        If REPROTENVIRONMENT.Connection.State = adStateOpen Then
            REPROTENVIRONMENT.Connection.Close
        End If
        REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
        REPROTENVIRONMENT.repFromCustDetail ModuleVarious.repSlnCust
        REPLACE_FROM_CUSTOMER.Show
    End If
 
    If capName = "Replace To Principal" Then
       odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
       If REPROTENVIRONMENT.Connection.State = adStateOpen Then
             REPROTENVIRONMENT.Connection.Close
       End If
       REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
       REPROTENVIRONMENT.repToPrnpalDetail ModuleVarious.repSlnPrnpal
       REP_TO_PRIN_DETAIL.Show
    End If
 
    If capName = "By Defective Product Entry Date" Then
        odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
        If REPROTENVIRONMENT.Connection.State = adStateOpen Then
               REPROTENVIRONMENT.Connection.Close
        End If
        REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
        REPROTENVIRONMENT.defProEDateDetail ModuleVarious.sDate, ModuleVarious.eDate
        DEFECTIVE_PRO_eDATE_DETAIL.Show
     End If
 
    If capName = "By Last Transaction Date" Then
        odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
        If REPROTENVIRONMENT.Connection.State = adStateOpen Then
            REPROTENVIRONMENT.Connection.Close
        End If
        REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
        REPROTENVIRONMENT.stkLastModifyDateQry ModuleVarious.sDate, ModuleVarious.eDate
        STOCK_LAST_TRAN_DATE_DETAIL.Show
     End If
 
    If capName = "By Date of Joining" Then
         odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
         If REPROTENVIRONMENT.Connection.State = adStateOpen Then
                REPROTENVIRONMENT.Connection.Close
         End If
         REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
         REPROTENVIRONMENT.crtEmpDOJQry ModuleVarious.sDate, ModuleVarious.eDate
         CRT_EMP_DOJ_DETAIL.Show
    End If
 
    If capName = "By Relieved Employee Date of Joining" Then
        odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
        If REPROTENVIRONMENT.Connection.State = adStateOpen Then
            REPROTENVIRONMENT.Connection.Close
        End If
        REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
        REPROTENVIRONMENT.rvdEmpDojQry ModuleVarious.sDate, ModuleVarious.eDate
        RELIEVED_EMP_DOJ_DETAIL.Show
    End If
 
    If capName = "By Relieved Employee Date of Relieved" Then
       odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
       If REPROTENVIRONMENT.Connection.State = adStateOpen Then
             REPROTENVIRONMENT.Connection.Close
       End If
       REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
       REPROTENVIRONMENT.rvdEmpDORQry ModuleVarious.sDate, ModuleVarious.eDate
       RELIEVED_EMP_DOR_DETAIL.Show
    End If
 
    If capName = "By Month of Payment" Then
        odbca = "driver={Microsoft ODBC for Oracle};uid=hrishi;pwd=jms;server=jms"
        If REPROTENVIRONMENT.Connection.State = adStateOpen Then
            REPROTENVIRONMENT.Connection.Close
        End If
        REPROTENVIRONMENT.Connection.Open odbca, "hrishi", "jms"
        REPROTENVIRONMENT.salEmpMonthQry ModuleVarious.monthN
        SALARY_EMP_MONTH_DETAIL.Show
    End If
 
 End Sub

Private Sub win_ArrangeIcons_Click()
    
    SANPROJECT.Arrange vbArrangeIcons

End Sub

Private Sub win_cascade_Click()

    SANPROJECT.Arrange vbCascade

End Sub

Private Sub win_tileHorizontally_Click()
    
    SANPROJECT.Arrange vbTileHorizontal

End Sub

Private Sub win_tileVerically_Click()

    SANPROJECT.Arrange vbTileVertical

End Sub

