Attribute VB_Name = "realize"
Public tools As New CITICSTools
Public database As New CITICSDatabase

' xRiskϵͳ���ݿ�������Ϣ
Public Const conxRiskzqInfo As String = "Provider=MSDAORA;" & _
                    "Data Source=xriskzx;" & _
                    "User Id=xriskzq;" & _
                    "Password=xriskzq;"

' ��Դ���ݿ�������Ϣ
Public Const conJYDBInfo As String = "Provider=MSDASQL;DSN=JYDB;UID=inforead;PWD=readinfo;"

' Wind
Public Const conWindInfo As String = "Provider=OraOLEDB.Oracle;" & _
                                "Data Source=winddata;" & _
                                "User Id=readwind;" & _
                                "Password=readwind;"

' xRiskϵͳ��д����Ȩ�޵�������Ϣ
Public Const conxRiskInfo As String = "Provider=OraOLEDB.Oracle;" & _
                                "Data Source=xriskzx;" & _
                                "User Id=xrisk;" & _
                                "Password=xrisk;PLSQLRSet=1;"
                               
' һ��Ԥ�����ݿ�������Ϣ
Public Const conGoGoalInfo As String = "Provider=MSDASQL;DSN=yzyq;UID=gogoal;PWD=gogoal;"
    '"Driver={SQL Server};Server=yjb_server_1;Database=FundRiskControl;Uid=gogoal;Pwd=gogoal;"

Public Const conxRisk33Info As String = ""
Public Const conGPJYInfo As String = ""

Sub Main()
    Debug.Print "dll initialize"
End Sub


