<script language="VBScript" RunAt="Server">
Function check_Pass()
  dim SQLstr,sa1,adocon,param 
  company = Request.Form("company")
   SQLstr = "select * from DownInfo where company='" & company & "'"
  '---------------------------------------------
  'Step 1�G�إ�ADO Connection�A�M��}��SQL��Access��Ʈw
   set adocon=Server.CreateObject("ADODB.Connection")
   adocon.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & server.MapPath("DownInfo")
  '---------------------------------------------
  'Step 2�G����SQL���O�A�ñN���G�x�s��Recordse��
  set sa1=adocon.Execute(SQLstr)
  '---------------------------------------------
  'Step 4�G����RecordSet�θ�Ʈw
  sa1.Close
  adocon.Close
End Function

Function adjust_sql(adj_str)
  '�ت��G�ഫ�d�ߦr�ꤤ���S���r����SQL���O
  Dim final_str, i
  '
  adj_str = Trim(adj_str)
  final_str = ""
  '�C���q�Ƕi�Ӫ��r�ꤤ���X�@�Ӧr���A�çP�_�O�_���U�z�T��
  '�r�����@�G�����A��"["�BDOS��Filter�Ÿ�"|"�H�γ�޸�"'"
  If Len(adj_str) > 0 Then
     For i = 1 To Len(adj_str)
       Select Case Mid(adj_str, i, 1)
          Case "[":
            final_str = final_str & "[[]"
          Case "|":
            final_str = final_str & "[{-}]"
          Case "'":
            final_str = final_str & "[&-(]"
          Case Else:  '�H�W�ҫD�h����Ƴ�����
            final_str = final_str & Mid(adj_str, i, 1)
       End Select
     Next
  End If
  adjust_sql = final_str
End Function

Sub Add_Reg()
     
   Dim company, name, idno, email, address
  
   Dim phone, fax,  sdate, chk_info
   Dim SQLstr
   '-----------------------------------------------
   On Error Resume Next
   '���X���������ܼƤ�
   company = Request.Form("company")
   name = Request.Form("name")
   address = Request.Form("address")
   no = Request.Form("no")
   phone = Request.Form("phone")
   fax = Request.Form("fax")
   'mail = Request.Form("mail")
      
   sdate = Request("syear") & Request("smonth") & Request("sday")
   '-----------------------------------------------
   '�վ�SQL���O�����S�w�r��
   company = adjust_sql(company)
  
   name = adjust_sql(name)
   no = adjust_sql(no)
   address = adjust_sql(address)
   phone = adjust_sql(phone)
   fax = adjust_sql(fax)
   '-----------------------------------------------
   SQLstr = "insert into DownInfo values ('" & company
   SQLstr = SQLstr & "','" & fax & "','" & no
   SQLstr = SQLstr & "','" & name & "','" & address & "','"
   SQLstr = SQLstr & "','" & phone & "','" & sdate & "')"
   '-----------------------------------------------
   'Step 1�G�إ�ADO Connection�A�M��}��Access��SQL��Ʈw
   set adocon=Server.CreateObject("ADODB.Connection")
   adocon.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & server.MapPath("DownInfo")
   '-----------------------------------------------
 
  
   '�Y�����~�o��
   If Err.Number <> 0 Then
      Response.Write "���~���X�G" & Err.Number & "<br>"
      Response.Write "���~�{���G" & Err.Source & "<br>"
      Response.Write "���~�T���G" & Err.Description & "<br>"
   Else  '�N���\�T����X���s������
      Response.Write("<hr><h1>")
      Response.Write("<center>���H�I<p>")
      Response.Write("���U���\�A�Ц^�e�@�e���I")
      Response.Write("</center>")
      Response.Write("</h1><hr>")
   End If
   '-----------------------------------------------
   'Step 3�G������Ʈw
   adocon.Close
   '-----------------------------------------------
End Sub


</script>
