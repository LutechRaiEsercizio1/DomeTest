Attribute VB_Name = "ACQ_Novita_m"
Option Explicit

Private Const HKEY_USERS = &H80000003
Private Const KEY_QUERY_VALUE = &H1
Private Const REG_SZ = 1                         ' Unicode nul terminated string

Const ITALIAN_DATE_FORMAT As String = "dd/mm/yyyy"
Const AMERICAN_DATE_FORMAT As String = "mm/dd/yyyy"

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Const INI_FILE_NAME = "acq_novita.ini"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public FileEsito As String
Dim NomeEsitoDapas As String
Dim ErroreHost As Integer

Dim ChkSalto As Boolean
Dim versione As Integer
Dim DataOraInizioMC As String
Dim EsitoMC As Boolean
Dim OccupataDaErrore As Boolean
Dim RichAcq As Integer
Dim AppIniFile As String    'prima era in ogni sub o func, ora e' globale
   
Dim conn As Object
Dim ORA_conn As Object
Private EnableLog As Boolean

Private Function FormatDate(NewDate As String) As String
   Static English_Style_Date As Integer
   
   Dim anno As String
   Dim mese As String
   Dim giorno As String
   
   If English_Style_Date = 0 Then
      'questo blocco viene eseguito una volta sola durante la quale
      'la variabile English_Style_Date potrà assumere:
      'il valore 1 se occorre la data di sistema viene espressa in termini mm/dd/yyyy
      'il valore -1 altrimenti
      If Mid(GetDateFormat, 1, 1) <> "D" Then
         English_Style_Date = 1
      Else
         English_Style_Date = -1
      End If
   End If
   
   anno = Format(NewDate, "YYYY")
   mese = Format(NewDate, "MM")
   giorno = Format(NewDate, "DD")
   If English_Style_Date = 1 And CInt(giorno) <= 12 Then Swap giorno, mese
   FormatDate = anno & mese & giorno
End Function

Private Function GetDateFormat() As String
   Dim subKey As String
   Dim options As Long
   Dim hdlReg As Long
   Dim key As String
   Dim N As Long
   Static DateFormatString As String
   
   If DateFormatString = "" Then
      subKey = ".DEFAULT\Control Panel\International"
      options = 0
      RegOpenKeyEx HKEY_USERS, subKey, options, KEY_QUERY_VALUE, hdlReg
      key = "sShortDate"
      DateFormatString = Space(255)
      N = Len(DateFormatString)
      RegQueryValueEx hdlReg, key, 0, REG_SZ, DateFormatString, N
      RegCloseKey hdlReg
      DateFormatString = UCase(Mid(DateFormatString, 1, N - 1))
   End If
   GetDateFormat = DateFormatString
End Function

Private Function GetServerCOMTI() As String
    Dim N As Integer
    Dim ret As Long
    Dim serverName As String
    
    serverName = Space(255)
    N = Len(serverName)
    ret = GetPrivateProfileString("comti", "server", "", serverName, N, AppIniFile)
    serverName = Mid(serverName, 1, ret)
    GetServerCOMTI = serverName
End Function

Private Function GetDBConnection() As String

   Const INI_DB_SECTION As String = "db_parameters"
   Const DB_DRIVER As String = "driver"
   Const DB_SERVER As String = "server"
   Const DB_PASSWORD As String = "pwd"
   Const DB_USERID As String = "uid"
   Const DB_INSTANCE As String = "database"

   Dim N As Integer
   Dim ret As Long
      
   Dim Driver As String
   Dim Server As String
   Dim Pwd As String
   Dim Uid As String
   Dim Instance As String
   
On Error GoTo err_getDBConn

   Driver = Space(255)
   N = Len(Driver)
   ret = GetPrivateProfileString(INI_DB_SECTION, DB_DRIVER, "", Driver, N, AppIniFile)
   Driver = Mid(Driver, 1, ret)
   
   Server = Space(255)
   N = Len(Server)
   ret = GetPrivateProfileString(INI_DB_SECTION, DB_SERVER, "", Server, N, AppIniFile)
   Server = Mid(Server, 1, ret)
   
   Pwd = Space(255)
   N = Len(Pwd)
   ret = GetPrivateProfileString(INI_DB_SECTION, DB_PASSWORD, "", Pwd, N, AppIniFile)
   Pwd = Mid(Pwd, 1, ret)
   
   Uid = Space(255)
   N = Len(Uid)
   ret = GetPrivateProfileString(INI_DB_SECTION, DB_USERID, "", Uid, N, AppIniFile)
   Uid = Mid(Uid, 1, ret)
   
   Instance = Space(255)
   N = Len(Instance)
   ret = GetPrivateProfileString(INI_DB_SECTION, DB_INSTANCE, "", Instance, N, AppIniFile)
   Instance = Mid(Instance, 1, ret)
   
    GetDBConnection = "driver=" & Driver & ";server=" & Server & ";pwd=" & Pwd & ";uid=" & Uid & ";database=" & Instance
    Exit Function
    
err_getDBConn:
    WriteLog "Errore nella funzione GetDBConnection : " & Err.Description
    GetDBConnection = ""
End Function
Private Function GetORADBConnection() As String

   Const INI_DB_SECTION As String = "dbora_parameters"
   Const DB_STRINGACONN As String = "stringa_conn"

   Dim N As Integer
   Dim ret As Long
   Dim stringa_conn As String
   Dim Instance As String
   
On Error GoTo err_getORADBConn

   stringa_conn = Space(255)
   N = Len(stringa_conn)
   ret = GetPrivateProfileString(INI_DB_SECTION, DB_STRINGACONN, "", stringa_conn, N, AppIniFile)
   stringa_conn = Mid(stringa_conn, 1, ret)
  
   GetORADBConnection = stringa_conn
   Exit Function
    
err_getORADBConn:
    WriteLog "Errore nella funzione GetORADBConnection : " & Err.Description
    GetORADBConnection = ""
End Function

Function getVersPrecRich(conn As Object, Arich, Nrich, VerRic) As Integer
'invece di cercare la versione immediatamente precedente, ricerca l'ultima
'versione che e' stata ottimizzata (quindi si deve accedere alla tabella delle risorse)
On Error GoTo err_handler
Dim sqlRich, rsRich

   sqlRich = "SELECT id_ver_Rich FROM Risorse " & _
      "WHERE num_ottim is not null and id_Anno_Rich = " & Arich & _
      " and  id_Num_Rich = " & Nrich & _
      " and  id_Ver_Rich < " & VerRic & " order by id_Ver_Rich desc"
   Set rsRich = conn.execute(sqlRich)
   If rsRich.EOF Then
      getVersPrecRich = 0
   Else
      getVersPrecRich = rsRich("id_ver_Rich")
   End If
   rsRich.Close
   Set rsRich = Nothing
   Exit Function
err_handler:
    getVersPrecRich = 0
    WriteLog "errore nella ricerca della versione precedente della richiesta : " & Err.Description
    WriteLog "SQL : " & sqlRich
End Function

Function isStessiDati_Mezzi(ByRef Rich As RaiBook.M7375350_RICHIESTE, cn As Object, ByVal maxProg As Long) As Boolean
Dim sql, rs, i, countEsistenti, countNuovi

    isStessiDati_Mezzi = False
    
    On Error GoTo err_handler
    
    With Rich
    sql = "Select count(*) From MezziRipresa Where " & _
    "id_anno_rich=" & .A7375351_ANNO_RICH & " and id_num_rich=" & .A7375351_NUM_RICH & _
    " and id_ver_rich=" & .A7375351_VER_RICH & " AND Versione_MR=" & maxProg
    
    'conteggio records
    Set rs = cn.execute(sql)
    countEsistenti = rs(0)
    rs.Close

    sql = "Select * From MezziRipresa Where " & _
    "id_anno_rich=" & .A7375351_ANNO_RICH & " and id_num_rich=" & .A7375351_NUM_RICH & _
    " and id_ver_rich=" & .A7375351_VER_RICH & " AND Versione_MR=" & maxProg
    
    Set rs = cn.execute(sql)
    If rs.EOF Then  'nuovi mezzi, deve scriverli comunque
        rs.Close
        Exit Function
    End If
    
    For i = LBound(Rich.A7375351_EL_MR_MEZZO) To UBound(Rich.A7375351_EL_MR_MEZZO)
'        countEsistenti = countEsistenti + 1
        If Trim(Rich.A7375351_EL_MR_MEZZO(i).A7375351_MR_MEZZO) = "" Then
            Exit For
        End If
        countNuovi = countNuovi + 1
        
        If Trim(Rich.A7375351_EL_MR_MEZZO(i).A7375351_MR_MEZZO) <> Trim(rs("mezzo_mr")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        If Trim(Rich.A7375351_EL_MR_MEZZO(i).A7375351_MR_DESCR) <> Trim(rs("descrizione_mr")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        If Trim(Rich.A7375351_EL_MR_MEZZO(i).A7375351_MR_ALTRO) <> Trim(rs("altro_mr")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        rs.movenext
        If rs.EOF Then
            Exit For
        End If
    Next i
    
    If countNuovi = countEsistenti And rs.EOF Then
        isStessiDati_Mezzi = True
        WriteLog "Mezzi di ripresa invariati"
    Else
        WriteLog "Mezzi nuovi " & countNuovi & " Mezzi esistenti " & countEsistenti & " eof=" & rs.EOF
    End If
    
    rs.Close
    
   End With
   Exit Function
   
err_handler:
     
    WriteLog "errore nella funzione di confronto dei mezzi di Collegamento : " & Err.Description
    WriteLog "SQL : " & sql
End Function

Function isStessiDati_CollegMobili(ByRef Rich As RaiBook.M7375350_RICHIESTE, cn As Object, ByVal maxProg As Long) As Boolean
Dim sql, rs, oraI_C1, oraI_C2, oraF_C1, oraF_C2, hh, mm

    isStessiDati_CollegMobili = False
    On Error GoTo err_handler
    With Rich
    
    sql = "Select * From CollegamentiMobili Where " & _
    "anno_rich_cm=" & .A7375351_ANNO_RICH & " and num_rich_cm=" & .A7375351_NUM_RICH & _
    " and ver_rich_cm=" & .A7375351_VER_RICH & " AND Versione_CM=" & maxProg
    
    Set rs = cn.execute(sql)
    If rs.EOF Then  'nuovi mezzi, deve scriverli comunque
        rs.Close
        Exit Function
    End If
    
    hh = CInt(Hour("0" & rs("C1_Ora_Inizio_CM")))
    mm = CInt(Minute("0" & rs("C1_Ora_Inizio_CM")))
    oraI_C1 = Right("00" & hh, 2) & Right("00" & mm, 2)
    
    hh = CInt(Hour("0" & rs("C2_Ora_Inizio_CM")))
    mm = CInt(Minute("0" & rs("C2_Ora_Inizio_CM")))
    oraI_C2 = Right("00" & hh, 2) & Right("00" & mm, 2)
    
    hh = CInt(Hour("0" & rs("C1_Ora_Fine_CM")))
    mm = CInt(Minute("0" & rs("C1_Ora_Fine_CM")))
    oraF_C1 = Right("00" & hh, 2) & Right("00" & mm, 2)
    
    hh = CInt(Hour("0" & rs("C2_Ora_Fine_CM")))
    mm = CInt(Minute("0" & rs("C2_Ora_Fine_CM")))
    oraF_C2 = Right("00" & hh, 2) & Right("00" & mm, 2)
        
    If CInt(.A7375351_CM_ORA_INIZ_PM) <> CInt(oraI_C1) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If CInt(.A7375351_CM_ORA_FINE_PM) <> CInt(oraF_C1) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If CInt(.A7375351_C2_ORA_INIZ_PM) <> CInt(oraI_C2) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If CInt(.A7375351_C2_ORA_FINE_PM) <> CInt(oraF_C2) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_CM_SISTEMA) <> Trim(rs("C1_Sistema_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_C2_SISTEMA) <> Trim(rs("C2_Sistema_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_CM_NUM_TELEFONO) <> Trim(rs("C1_Telefono_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_C2_NUM_TELEFONO) <> Trim(rs("C2_Telefono_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_CM_LOCALITA_RX) <> Trim(rs("C1_Localita_RX_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_C2_LOCALITA_RX) <> Trim(rs("C2_Localita_RX_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_CM_NOME_TXP) <> Trim(rs("CM_CANALE")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_C2_NOME_TXP) <> Trim(rs("C2_CANALE")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_CM_AUDIO1) <> Trim(rs("C1_Audio1_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_CM_AUDIO2) <> Trim(rs("C1_Audio2_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
     If Trim(Rich.A7375351_C2_AUDIO1) <> Trim(rs("C2_Audio1_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_C2_AUDIO2) <> Trim(rs("C2_Audio2_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_CM_NOTE) <> Trim(rs("C1_NOTE_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
    If Trim(Rich.A7375351_C2_NOTE) <> Trim(rs("C2_NOTE_CM")) Then
        rs.Close    'dati cambiati
        Exit Function
    End If
   
    isStessiDati_CollegMobili = True
    
    rs.Close
    
    End With
    WriteLog "Collegamenti mobili invariati"
    Exit Function
   
err_handler:
    WriteLog "errore nella funzione di confronto dei Ponti Mobili : " & Err.Description
    WriteLog "SQL : " & sql
End Function
Private Function isStessiDati_NewCollegMobili(cn As Object, ByVal ANNO_RICHIESTA As Integer, ByVal NUM_RICHIESTA As Long, ByVal ver_richiesta As Integer, ByVal maxProg As Long, rstCM1 As Object, no_C2 As Boolean, Optional rstCM2 As Object) As Boolean
Dim sql, rs, oraI_C1, oraI_C2, oraF_C1, oraF_C2, hh, mm

    isStessiDati_NewCollegMobili = False
    
    On Error GoTo err_handler
    
    Dim CM_Ora_Iniz_PM As String
    Dim CM_Ora_Fine_PM As String
    Dim WS_Ora_Iniz_PM As String
    Dim WS_Ora_Fine_PM As String
    Dim cod_satellite As String
    Dim nro_cod_satellite As String
    Dim d() As String
       
    sql = "Select * From CollegamentiMobili Where " & _
    "anno_rich_cm=" & ANNO_RICHIESTA & " and num_rich_cm=" & NUM_RICHIESTA & _
    " and ver_rich_cm=" & ver_richiesta & " AND Versione_CM=" & maxProg
    
    Set rs = cn.execute(sql)
    If rs.EOF Then
        rs.Close
        Exit Function
    End If
     
    If rstCM1.EOF Then
        WriteLog "non trovato il recordset WinScroll"
        Exit Function
    End If
    'C1
    
   If IsNull(rs("C1_Ora_Inizio_CM")) Then
        CM_Ora_Iniz_PM = "'1900/01/01 00:00:00'"
   Else
        CM_Ora_Iniz_PM = "'1900/01/01 " & Right("00" & Hour(rs("C1_Ora_Inizio_CM")), 2) & ":" & Right("00" & Minute(rs("C1_Ora_Inizio_CM")), 2) & ":" & Right("00" & Second(rs("C1_Ora_Inizio_CM")), 2) & "'"
   End If
        
    
   If IsNull(rs("C1_Ora_Fine_CM")) Then
        CM_Ora_Fine_PM = "'1900/01/01 00:00:00'"
   Else
        CM_Ora_Fine_PM = "'1900/01/01 " & Right("00" & Hour(rs("C1_Ora_Fine_CM")), 2) & ":" & Right("00" & Minute(rs("C1_Ora_Fine_CM")), 2) & ":" & Right("00" & Second(rs("C1_Ora_Fine_CM")), 2) & "'"
   End If

     
   If IsNull(rstCM1("Ora_Inizio_Pian")) Then
        WS_Ora_Iniz_PM = "'1900/01/01 00:00:00'"
   Else
        WS_Ora_Iniz_PM = "'1900/01/01 " & Right("00" & Hour(rstCM1("Ora_Inizio_Pian")), 2) & ":" & Right("00" & Minute(rstCM1("Ora_Inizio_Pian")), 2) & ":" & Right("00" & Second(rstCM1("Ora_Inizio_Pian")), 2) & "'"
   End If
        
    
   If IsNull(rstCM1("Ora_Fine_Pian")) Then
        WS_Ora_Fine_PM = "'1900/01/01 00:00:00'"
   Else
        WS_Ora_Fine_PM = "'1900/01/01 " & Right("00" & Hour(rstCM1("Ora_Fine_Pian")), 2) & ":" & Right("00" & Minute(rstCM1("Ora_Fine_Pian")), 2) & ":" & Right("00" & Second(rstCM1("Ora_Fine_Pian")), 2) & "'"
   End If
   
     
   'WriteLog "C1_Ora_Inizio_CM : " & CM_Ora_Iniz_PM
   'WriteLog "C1_Ora_Fine_CM : " & CM_Ora_Fine_PM
   'WriteLog "Ora_Inizio_WS : " & WS_Ora_Iniz_PM
   'WriteLog "Ora_Fine_WS : " & WS_Ora_Fine_PM
    
      
    If Trim(CM_Ora_Iniz_PM) <> Trim(WS_Ora_Iniz_PM) Then
        WriteLog "Date diverse : " & CM_Ora_Iniz_PM & " - " & WS_Ora_Iniz_PM
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    If Trim(CM_Ora_Fine_PM) <> Trim(WS_Ora_Fine_PM) Then
        WriteLog "Date diverse : " & CM_Ora_Iniz_PM & " - " & WS_Ora_Iniz_PM
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    If Trim(rstCM1("Mezzo_trasmissione")) <> Trim(rs("C1_Sistema_CM")) Then
        WriteLog "Mezzo Trasmissione"
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    If Trim(rstCM1("Telefono_Mezzo_tx")) <> Trim(rs("C1_Telefono_CM")) Then
        WriteLog "Telefono"
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    If Trim(rstCM1("Transponder")) <> Trim(rs("C1_NOME_TXP_CM")) Then
        WriteLog "Txp"
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    If Trim(rstCM1("Audio_stereo1")) <> Trim(rs("C1_Audio1_CM")) Then
        WriteLog "a1"
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    If Trim(rstCM1("Audio_stereo2")) <> Trim(rs("C1_Audio2_CM")) Then
        WriteLog "a2"
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    If Trim(rstCM1("NOTE_output")) <> Trim(rs("C1_NOTE_CM")) Then
        WriteLog "Note"
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    'NUOVI CAMPI
    
If Trim(rstCM1("Origine_Collegamento")) <> Trim(rs("CM_Origine")) Then
    rs.Close    'dati cambiati
        WriteLog "Origine"
    Exit Function
End If
   
If Trim(rstCM1("Destinazione_Collegamento")) <> Trim(rs("CM_Destinazione")) Then
    rs.Close    'dati cambiati
        WriteLog "Destinazione"
    Exit Function
End If
  
 If Trim(rstCM1("Mezzo_ricezione")) <> Trim(rs("C1_Mezzo_Ricez_CM")) Then
    rs.Close    'dati cambiati
        WriteLog "Ricez"
    Exit Function
End If
    
If Trim(rstCM1("Ind_Miaf")) <> Trim(rs("C1_Miaf_CM")) Then
    rs.Close    'dati cambiati
        WriteLog "Miaf"
    Exit Function
End If
    
    
If Trim(rstCM1("Satellite")) <> Trim(rs("CM_Satellite")) Then
        WriteLog "Sat"
    rs.Close    'dati cambiati
    Exit Function
End If

If Trim(rstCM1("Canale")) <> Trim(rs("CM_Canale")) Then
        WriteLog "Can"
    rs.Close    'dati cambiati
    Exit Function
End If

If Trim(rstCM1("Frequenza")) <> Trim(rs("CM_Frequenza")) Then
        WriteLog "freq"
    rs.Close    'dati cambiati
    Exit Function
End If
    
If Trim(rstCM1("Polarizzazione")) <> Trim(rs("CM_Polarizzazione")) Then
        WriteLog "pol"
    rs.Close    'dati cambiati
    Exit Function
End If
    
If Trim(rstCM1("Posizione_orbitale")) <> Trim(rs("CM_Posizione_orbitale")) Then
        WriteLog "pos"
    rs.Close    'dati cambiati
    Exit Function
End If
    
If Trim(rstCM1("Sistema_tx_sat")) <> Trim(rs("CM_System")) Then
        WriteLog "sys"
    rs.Close    'dati cambiati
    Exit Function
End If
    
If Trim(rstCM1("Multiplex_tx_type")) <> Trim(rs("CM_MTT")) Then
        WriteLog "mtt"
    rs.Close    'dati cambiati
    Exit Function
End If
    
If Trim(rstCM1("FEC")) <> Trim(rs("CM_FEC")) Then
        WriteLog "fec"
    rs.Close    'dati cambiati
    Exit Function
End If
    
If Trim(rstCM1("Bitrate")) <> Trim(rs("CM_Bit_rate")) Then
        WriteLog "br"
    rs.Close    'dati cambiati
    Exit Function
End If
   
If Trim(rstCM1("Symbolrate")) <> Trim(rs("CM_Symbol_rate")) Then
        WriteLog "sr"
    rs.Close    'dati cambiati
    Exit Function
End If
   
If Trim(rstCM1("Banda")) <> Trim(rs("CM_Banda")) Then
        WriteLog "banda"
    rs.Close    'dati cambiati
    Exit Function
End If


' M.G. nuovi atributi satellitari e codiceservizio
If Trim(rstCM1("frequenzadown")) <> Trim(rs("CM_Freq_Dw_link")) Then
    rs.Close    'dati cambiati
    Exit Function
End If
   
If Trim(rstCM1("polarizzazionedown")) <> Trim(rs("CM_Pol_Dw_link")) Then
    rs.Close    'dati cambiati
    Exit Function
End If
   
If Trim(rstCM1("codiceservizio")) <> Trim(rs("CM_Codice_Servizio")) Then
    rs.Close    'dati cambiati
    Exit Function
End If



nro_cod_satellite = Replace(rstCM1("Codif_Sat") & "", "'", "''")

If Trim(rstCM1("Codif_Sat")) <> "" Then
    If InStr(rstCM1("Codif_Sat"), "-") > 0 Then
        d = Split(nro_cod_satellite, "-")
            If LBound(d) = 0 Then
                cod_satellite = d(0)
            End If
            If UBound(d) > 0 Then
                nro_cod_satellite = d(1)
            End If
    
        If Trim(cod_satellite) <> Trim(rs("CM_Cod_Satellite")) Then
            WriteLog "cod sat : " & cod_satellite & " - " & Trim(rs("CM_Cod_Satellite"))
            rs.Close    'dati cambiati
            Exit Function
        End If
        If Trim(nro_cod_satellite) <> Trim(rs("CM_NRO_Cod_Satellite")) Then
            rs.Close    'dati cambiati
            WriteLog "nro cod sat"
            Exit Function
        End If
    End If
End If
   
'------------------------------------------------------------------------------
   
    'C2
    If no_C2 Then 'bypassare i test su c2
        isStessiDati_NewCollegMobili = True
        WriteLog "Collegamenti mobili invariati"
        rs.Close
        Exit Function
    End If
    If rstCM2.EOF Then
        isStessiDati_NewCollegMobili = True
        WriteLog "Collegamenti mobili invariati"
        rs.Close
        Exit Function
    End If
        
   If IsNull(rs("C2_Ora_Inizio_CM")) Then
        CM_Ora_Iniz_PM = "'1900/01/01 00:00:00'"
   Else
        CM_Ora_Iniz_PM = "'1900/01/01 " & Right("00" & Hour(rs("C2_Ora_Inizio_CM")), 2) & ":" & Right("00" & Minute(rs("C2_Ora_Inizio_CM")), 2) & ":" & Right("00" & Second(rs("C2_Ora_Inizio_CM")), 2) & "'"
   End If
        
    
   If IsNull(rs("C2_Ora_Fine_CM")) Then
        CM_Ora_Fine_PM = "'1900/01/01 00:00:00'"
   Else
        CM_Ora_Fine_PM = "'1900/01/01 " & Right("00" & Hour(rs("C2_Ora_Fine_CM")), 2) & ":" & Right("00" & Minute(rs("C2_Ora_Fine_CM")), 2) & ":" & Right("00" & Second(rs("C2_Ora_Fine_CM")), 2) & "'"
   End If

     
   If IsNull(rstCM2("Ora_Inizio_Pian")) Then
        WS_Ora_Iniz_PM = "'1900/01/01 00:00:00'"
   Else
        WS_Ora_Iniz_PM = "'1900/01/01 " & Right("00" & Hour(rstCM2("Ora_Inizio_Pian")), 2) & ":" & Right("00" & Minute(rstCM2("Ora_Inizio_Pian")), 2) & ":" & Right("00" & Second(rstCM2("Ora_Inizio_Pian")), 2) & "'"
   End If
        
    
   If IsNull(rstCM2("Ora_Fine_Pian")) Then
        WS_Ora_Fine_PM = "'1900/01/01 00:00:00'"
   Else
        WS_Ora_Fine_PM = "'1900/01/01 " & Right("00" & Hour(rstCM2("Ora_Fine_Pian")), 2) & ":" & Right("00" & Minute(rstCM2("Ora_Fine_Pian")), 2) & ":" & Right("00" & Second(rstCM2("Ora_Fine_Pian")), 2) & "'"
   End If
   
     
   'WriteLog "C1_Ora_Inizio_CM : " & CM_Ora_Iniz_PM
   'WriteLog "C1_Ora_Fine_CM : " & CM_Ora_Fine_PM
   'WriteLog "Ora_Inizio_WS : " & WS_Ora_Iniz_PM
   'WriteLog "Ora_Fine_WS : " & WS_Ora_Fine_PM
    
      
    If Trim(CM_Ora_Iniz_PM) <> Trim(WS_Ora_Iniz_PM) Then
        WriteLog "Date diverse : " & CM_Ora_Iniz_PM & " - " & WS_Ora_Iniz_PM
        rs.Close    'dati cambiati
        Exit Function
    End If
    
    If Trim(CM_Ora_Fine_PM) <> Trim(WS_Ora_Fine_PM) Then
        WriteLog "Date diverse : " & CM_Ora_Iniz_PM & " - " & WS_Ora_Iniz_PM
        rs.Close    'dati cambiati
        Exit Function
    End If
        
        
        If Trim(rstCM2("Mezzo_trasmissione")) <> Trim(rs("C2_Sistema_CM")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        If Trim(rstCM2("Telefono_mezzo_tx")) <> Trim(rs("C2_Telefono_CM")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        If Trim(rstCM2("Transponder")) <> Trim(rs("C2_NOME_TXP_CM")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        If Trim(rstCM2("NOTE_output")) <> Trim(rs("C2_NOTE_CM")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
         If Trim(rstCM2("Audio_stereo1")) <> Trim(rs("C2_Audio1_CM")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        If Trim(rstCM2("Audio_stereo2")) <> Trim(rs("C2_Audio2_CM")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        
        
        
    'NUOVI CAMPI
    
        If Trim(rstCM2("Origine_Collegamento")) <> Trim(rs("C2_Origine")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
           
        If Trim(rstCM2("Destinazione_Collegamento")) <> Trim(rs("C2_Destinazione")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
          
         If Trim(rstCM2("Mezzo_ricezione")) <> Trim(rs("C2_Mezzo_Ricez_CM")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
            
        If Trim(rstCM2("Ind_Miaf")) <> Trim(rs("C2_Miaf_CM")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
            
            
        If Trim(rstCM2("Satellite")) <> Trim(rs("C2_Satellite")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        
        If Trim(rstCM2("Canale")) <> Trim(rs("C2_Canale")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        
        If Trim(rstCM2("Frequenza")) <> Trim(rs("C2_Frequenza")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
            
        If Trim(rstCM2("Polarizzazione")) <> Trim(rs("C2_Polarizzazione")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
            
        If Trim(rstCM2("Posizione_orbitale")) <> Trim(rs("C2_Posizione_orbitale")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
            
        If Trim(rstCM2("Sistema_tx_sat")) <> Trim(rs("C2_System")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
            
        If Trim(rstCM2("Multiplex_tx_type")) <> Trim(rs("C2_MTT")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
            
        If Trim(rstCM2("FEC")) <> Trim(rs("C2_FEC")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
            
        If Trim(rstCM2("Bitrate")) <> Trim(rs("C2_Bit_rate")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
           
        If Trim(rstCM2("Symbolrate")) <> Trim(rs("C2_Symbol_rate")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
           
        If Trim(rstCM2("Banda")) <> Trim(rs("C2_Banda")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        
        ' M.G. nuovi atributi satellitari e codiceservizio
        If Trim(rstCM2("frequenzadown")) <> Trim(rs("C2_Freq_Dw_link")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
   
        If Trim(rstCM2("polarizzazionedown")) <> Trim(rs("C2_Pol_Dw_link")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
   
        If Trim(rstCM2("codiceservizio")) <> Trim(rs("C2_Codice_Servizio")) Then
            rs.Close    'dati cambiati
            Exit Function
        End If
        
nro_cod_satellite = Replace(rstCM2("Codif_Sat") & "", "'", "''")
     If Trim(rstCM2("Codif_Sat")) <> "" Then
         If InStr(rstCM2("Codif_Sat"), "-") > 0 Then
             d = Split(nro_cod_satellite, "-")
                If LBound(d) = 0 Then
                    cod_satellite = d(0)
                End If
                If UBound(d) > 0 Then
                    nro_cod_satellite = d(1)
                End If
         End If
         If cod_satellite <> Trim(rs("C2_Cod_Satellite")) Then
            rs.Close    'dati cambiati
             Exit Function
         End If
         If nro_cod_satellite <> Trim(rs("C2_NRO_Cod_Satellite")) Then
             rs.Close    'dati cambiati
             Exit Function
         End If
     End If
        
    isStessiDati_NewCollegMobili = True
    
    rs.Close
    
    WriteLog "Collegamenti mobili invariati"
    Exit Function
err_handler:
    WriteLog "errore nella funzione di confronto dei Ponti Mobili : " & Err.Description
    WriteLog "SQL : " & sql
    Err.Clear
    End Function
Private Function getInsertCollegamentoMobile(ByVal ver As Integer, ByRef Richiesta As RaiBook.M7375350_RICHIESTE) As String
With Richiesta

Dim ANNO_RICHIESTA As String
Dim NUM_RICHIESTA As String

ANNO_RICHIESTA = .A7375351_ANNO_RICH
NUM_RICHIESTA = .A7375351_NUM_RICH

On Error GoTo err_getInsertCollegamentoMobile

Dim CM_Ora_Iniz_PM As String
Dim CM_Ora_Fine_PM As String
Dim C2_Ora_Iniz_PM As String
Dim C2_Ora_Fine_PM As String

Dim sql As String, txp1, txp2, loc1, loc2
Dim sist1, sist2, note1, note2, audio1_1, audio1_2, audio2_1, audio2_2, tel1, tel2

Dim satellite1 As String
Dim Cod_satellite1 As String
Dim NRO_Cod_satellite1 As String
Dim Posizione_orbitale1 As String
Dim Canale1 As String
Dim System1 As String
Dim Frequenza1 As String
Dim Polarizzazione1 As String
Dim Bit_Rate1 As String
Dim Symbol_Rate1 As String
Dim MTT1 As String
Dim FEC1 As String
Dim Origine1 As String
Dim Destinazione1 As String
Dim Miaf1 As String
Dim Mezzo_ricez1 As String
Dim Banda1 As String

Dim satellite2 As String
Dim Cod_satellite2 As String
Dim NRO_Cod_satellite2 As String
Dim Posizione_orbitale2 As String
Dim Canale2 As String
Dim System2 As String
Dim Frequenza2 As String
Dim Polarizzazione2 As String
Dim Bit_Rate2 As String
Dim Symbol_Rate2 As String
Dim MTT2 As String
Dim FEC2 As String
Dim Origine2 As String
Dim Destinazione2 As String
Dim Miaf2 As String
Dim Mezzo_ricez2 As String
Dim Banda2 As String

' M.G. Nuovi CodiceServizio e attributi satellitari
Dim CM_Freq_Dw_link As String
Dim CM_Pol_Dw_link As String
Dim CM_Codice_Servizio As String
Dim C2_Freq_Dw_link As String
Dim C2_Pol_Dw_link As String
Dim C2_Codice_Servizio As String
' M.G. -----------------------------------------

satellite1 = ""
Cod_satellite1 = ""
NRO_Cod_satellite1 = ""
Posizione_orbitale1 = ""
Canale1 = ""
System1 = ""
Frequenza1 = ""
Polarizzazione1 = ""
Bit_Rate1 = ""
Symbol_Rate1 = ""
MTT1 = ""
FEC1 = ""
Origine1 = ""
Destinazione1 = ""
Miaf1 = ""
Mezzo_ricez1 = ""
Banda1 = ""

satellite2 = ""
Cod_satellite2 = ""
NRO_Cod_satellite2 = ""
Posizione_orbitale2 = ""
Canale2 = ""
System2 = ""
Frequenza2 = ""
Polarizzazione2 = ""
Bit_Rate2 = ""
Symbol_Rate2 = ""
MTT2 = ""
FEC2 = ""
Origine2 = ""
Destinazione2 = ""
Miaf2 = ""
Mezzo_ricez2 = ""
Banda2 = ""

CM_Freq_Dw_link = ""
CM_Pol_Dw_link = ""
CM_Codice_Servizio = ""
C2_Freq_Dw_link = ""
C2_Pol_Dw_link = ""
C2_Codice_Servizio = ""


'collegamenti mobili 1
    If .A7375351_CM_ORA_INIZ_PM = 0 Then
       CM_Ora_Iniz_PM = "'1900/01/01 00:00.00'"
    Else
       CM_Ora_Iniz_PM = Right(String(4, "0") & .A7375351_CM_ORA_INIZ_PM, 4)
       CM_Ora_Iniz_PM = "'" & convertiOre24(Mid(CM_Ora_Iniz_PM, 1, 2) & ":" & Mid(CM_Ora_Iniz_PM, 3)) & "'"
    End If
    
    If .A7375351_CM_ORA_FINE_PM = 0 Then
       CM_Ora_Fine_PM = "'1900/01/01 00:00.00'"
    Else
       CM_Ora_Fine_PM = Right(String(4, "0") & .A7375351_CM_ORA_FINE_PM, 4)
       CM_Ora_Fine_PM = "'" & convertiOre24(Mid(CM_Ora_Fine_PM, 1, 2) & ":" & Mid(CM_Ora_Fine_PM, 3)) & "'"
    End If
      
'collegamenti mobili 2
    If .A7375351_C2_ORA_INIZ_PM = 0 Then
       C2_Ora_Iniz_PM = "'1900/01/01 00:00.00'"
    Else
       C2_Ora_Iniz_PM = Right(String(4, "0") & .A7375351_C2_ORA_INIZ_PM, 4)
       C2_Ora_Iniz_PM = "'" & convertiOre24(Mid(C2_Ora_Iniz_PM, 1, 2) & ":" & Mid(C2_Ora_Iniz_PM, 3)) & "'"
    End If
      
    If .A7375351_C2_ORA_FINE_PM = 0 Then
       C2_Ora_Fine_PM = "'1900/01/01 00:00.00'"
    Else
       C2_Ora_Fine_PM = Right(String(4, "0") & .A7375351_C2_ORA_FINE_PM, 4)
       C2_Ora_Fine_PM = "'" & convertiOre24(Mid(C2_Ora_Fine_PM, 1, 2) & ":" & Mid(C2_Ora_Fine_PM, 3)) & "'"
    End If
    sist1 = Replace(.A7375351_CM_SISTEMA, "'", "''")
    sist2 = Replace(.A7375351_C2_SISTEMA, "'", "''")
    tel1 = Replace(.A7375351_CM_NUM_TELEFONO, "'", "''")
    tel2 = Replace(.A7375351_C2_NUM_TELEFONO, "'", "''")
    loc1 = Replace(.A7375351_CM_LOCALITA_RX, "'", "''")
    loc2 = Replace(.A7375351_C2_LOCALITA_RX, "'", "''")
    Destinazione1 = Replace(.A7375351_CM_LOCALITA_RX, "'", "''")
    Destinazione2 = Replace(.A7375351_C2_LOCALITA_RX, "'", "''")
    Canale1 = Replace(.A7375351_CM_NOME_TXP, "'", "''")
    Canale2 = Replace(.A7375351_C2_NOME_TXP, "'", "''")
    txp1 = ""
    txp2 = ""
    audio1_1 = Replace(.A7375351_CM_AUDIO1, "'", "''")
    audio2_1 = Replace(.A7375351_CM_AUDIO2, "'", "''")
    audio1_2 = Replace(.A7375351_C2_AUDIO1, "'", "''")
    audio2_2 = Replace(.A7375351_C2_AUDIO2, "'", "''")
    note1 = Replace(.A7375351_CM_NOTE & "", "'", "''")
    note2 = Replace(.A7375351_C2_NOTE & "", "'", "''")
    
    sql = "INSERT INTO CollegamentiMobili (" & _
        "Anno_Rich_CM, Num_Rich_CM, Ver_Rich_CM, Versione_CM," & _
        "CM_satellite, CM_Cod_satellite, CM_NRO_Cod_satellite ,CM_Posizione_orbitale, " & _
        "CM_Canale, CM_System, CM_Frequenza, CM_Polarizzazione, " & _
        "CM_Bit_Rate, CM_Symbol_Rate, CM_MTT , CM_FEC, " & _
        "CM_Origine, CM_Destinazione, C1_Miaf_CM, C1_Mezzo_ricez_CM , CM_Banda , " & _
        "C2_satellite, C2_Cod_satellite, C2_NRO_Cod_satellite ,C2_Posizione_orbitale, " & _
        "C2_Canale, C2_System, C2_Frequenza, C2_Polarizzazione, " & _
        "C2_Bit_Rate, C2_Symbol_Rate, C2_MTT , C2_FEC, " & _
        "C2_Origine, C2_Destinazione, C2_Miaf_CM, C2_Mezzo_ricez_CM , C2_Banda , " & _
        "C1_Sistema_CM, C1_Telefono_CM, C1_Nome_TXP_CM, C1_Localita_RX_CM," & "C1_Audio1_CM, C1_Audio2_CM, C1_Ora_Inizio_CM, C1_Ora_Fine_CM, C1_Note_CM," & "C2_Sistema_CM, C2_Telefono_CM, C2_Nome_TXP_CM, C2_Localita_RX_CM," & "C2_Audio1_CM, C2_Audio2_CM, C2_Ora_Inizio_CM, C2_Ora_Fine_CM, C2_Note_CM,CM_Freq_Dw_link, CM_Pol_Dw_link, CM_Codice_Servizio, C2_Freq_Dw_link, C2_Pol_Dw_link, C2_Codice_Servizio" & _
        ") VALUES (" & _
            .A7375351_ANNO_RICH & ", " & .A7375351_NUM_RICH & ", " & .A7375351_VER_RICH & ", " & ver & " , " & _
            "'" & satellite1 & "', '" & Cod_satellite1 & "', " & "'" & NRO_Cod_satellite1 & "', '" & Posizione_orbitale1 & "', " & _
            "'" & Canale1 & "', '" & System1 & "', " & "'" & Frequenza1 & "', '" & Polarizzazione1 & "', " & _
            "'" & Bit_Rate1 & "', '" & Symbol_Rate1 & "', " & "'" & MTT1 & "', '" & FEC1 & "', " & _
            "'" & Origine1 & "', '" & Destinazione1 & "', " & "'" & Miaf1 & "', '" & Mezzo_ricez1 & "', '" & Banda1 & "', " & _
            "'" & satellite2 & "', '" & Cod_satellite2 & "', " & "'" & NRO_Cod_satellite2 & "', '" & Posizione_orbitale2 & "', " & _
            "'" & Canale2 & "', '" & System2 & "', " & "'" & Frequenza2 & "', '" & Polarizzazione2 & "', " & _
            "'" & Bit_Rate2 & "', '" & Symbol_Rate2 & "', " & "'" & MTT2 & "', '" & FEC2 & "', " & _
            "'" & Origine2 & "', '" & Destinazione2 & "', " & "'" & Miaf2 & "', '" & Mezzo_ricez2 & "', '" & Banda2 & "', " & _
            "'" & sist1 & "', '" & tel1 & "', " & "'" & txp1 & "', '" & loc1 & "', " & "'" & audio1_1 & "', '" & audio2_1 & "', " & CM_Ora_Iniz_PM & ", " & CM_Ora_Fine_PM & ", " & "'" & note1 & "', " & "'" & sist2 & "', '" & tel2 & "', " & _
            "'" & txp2 & "', '" & loc2 & "', " & "'" & audio1_2 & "', '" & audio2_2 & "', " & C2_Ora_Iniz_PM & ", " & C2_Ora_Fine_PM & ", '" & note2 & "', '" & CM_Freq_Dw_link & "', '" & CM_Pol_Dw_link & "', '" & CM_Codice_Servizio & "', '" & C2_Freq_Dw_link & "', '" & C2_Pol_Dw_link & "', '" & C2_Codice_Servizio & "')"
    
    End With

    getInsertCollegamentoMobile = sql
    Exit Function
err_getInsertCollegamentoMobile:
    getInsertCollegamentoMobile = ""
    WriteLog "Errore nella funzione di preparazione stringa di inserimento CM (vecchia) : " & Err.Description & " richiesta : " & ANNO_RICHIESTA & NUM_RICHIESTA
End Function

Private Function getInsertNewCollegamentoMobile(ByVal ANNO_RICHIESTA As Integer, ByVal NUM_RICHIESTA As Long, ByVal ver_richiesta As Integer, ByVal ver As Long, rstCM1 As Object, Optional rstCM2 As Object) As String

On Error GoTo err_getInsertNewCollegamentoMobile

Dim CM_Ora_Iniz_PM As String
Dim CM_Ora_Fine_PM As String
Dim C2_Ora_Iniz_PM As String
Dim C2_Ora_Fine_PM As String

Dim sql As String, txp1, txp2, loc1, loc2
Dim sist1, sist2, note1, note2, audio1_1, audio1_2, audio2_1, audio2_2, tel1, tel2
CM_Ora_Iniz_PM = "1900/01/01 00:00.00"
CM_Ora_Fine_PM = "1900/01/01 00:00.00"

Dim satellite1 As String
Dim Cod_satellite1 As String
Dim NRO_Cod_satellite1 As String
Dim Posizione_orbitale1 As String
Dim Canale1 As String
Dim System1 As String
Dim Frequenza1 As String
Dim Polarizzazione1 As String
Dim Bit_Rate1 As String
Dim Symbol_Rate1 As String
Dim MTT1 As String
Dim FEC1 As String
Dim Origine1 As String
Dim Destinazione1 As String
Dim Miaf1 As String
Dim Mezzo_ricez1 As String
Dim Banda1 As String

Dim satellite2 As String
Dim Cod_satellite2 As String
Dim NRO_Cod_satellite2 As String
Dim Posizione_orbitale2 As String
Dim Canale2 As String
Dim System2 As String
Dim Frequenza2 As String
Dim Polarizzazione2 As String
Dim Bit_Rate2 As String
Dim Symbol_Rate2 As String
Dim MTT2 As String
Dim FEC2 As String
Dim Origine2 As String
Dim Destinazione2 As String
Dim Miaf2 As String
Dim Mezzo_ricez2 As String
Dim Banda2 As String
Dim d() As String

' M.G. Nuovi CodiceServizio e attributi satellitari
Dim CM_Freq_Dw_link As String
Dim CM_Pol_Dw_link As String
Dim CM_Codice_Servizio As String
Dim C2_Freq_Dw_link As String
Dim C2_Pol_Dw_link As String
Dim C2_Codice_Servizio As String
' M.G. -----------------------------------------


satellite1 = ""
Cod_satellite1 = ""
NRO_Cod_satellite1 = ""
Posizione_orbitale1 = ""
Canale1 = ""
System1 = ""
Frequenza1 = ""
Polarizzazione1 = ""
Bit_Rate1 = ""
Symbol_Rate1 = ""
MTT1 = ""
FEC1 = ""
Origine1 = ""
Destinazione1 = ""
Miaf1 = ""
Mezzo_ricez1 = ""
Banda1 = ""

satellite2 = ""
Cod_satellite2 = ""
NRO_Cod_satellite2 = ""
Posizione_orbitale2 = ""
Canale2 = ""
System2 = ""
Frequenza2 = ""
Polarizzazione2 = ""
Bit_Rate2 = ""
Symbol_Rate2 = ""
MTT2 = ""
FEC2 = ""
Origine2 = ""
Destinazione2 = ""
Miaf2 = ""
Mezzo_ricez2 = ""
Banda2 = ""


sist1 = ""
tel1 = ""
loc1 = ""
txp1 = ""
audio1_1 = ""
audio2_1 = ""
note1 = ""
C2_Ora_Iniz_PM = "1900/01/01 00:00.00"
C2_Ora_Fine_PM = "1900/01/01 00:00.00"
sist2 = ""
tel2 = ""
loc2 = ""
txp2 = ""
audio1_2 = ""
audio2_2 = ""
note2 = ""

CM_Freq_Dw_link = ""
CM_Pol_Dw_link = ""
CM_Codice_Servizio = ""
C2_Freq_Dw_link = ""
C2_Pol_Dw_link = ""
C2_Codice_Servizio = ""
 
If Not rstCM1.EOF Then
'collegamenti mobili C1
    'CM_Ora_Iniz_PM = rstCM1("Ora_Inizio_Pian") & ""
    'CM_Ora_Fine_PM = rstCM1("Ora_Fine_Pian") & ""
    
   If IsNull(rstCM1("Ora_Inizio_Pian")) Then
        CM_Ora_Iniz_PM = "1900/01/01 00:00:00"
   Else
        CM_Ora_Iniz_PM = "1900/01/01 " & Right("00" & Hour(rstCM1("Ora_Inizio_Pian")), 2) & ":" & Right("00" & Minute(rstCM1("Ora_Inizio_Pian")), 2) & ":" & Right("00" & Second(rstCM1("Ora_Inizio_Pian")), 2)
   End If
        
    
   If IsNull(rstCM1("Ora_Fine_Pian")) Then
        CM_Ora_Fine_PM = "1900/01/01 00:00:00"
   Else
        CM_Ora_Fine_PM = "1900/01/01 " & Right("00" & Hour(rstCM1("Ora_Fine_Pian")), 2) & ":" & Right("00" & Minute(rstCM1("Ora_Fine_Pian")), 2) & ":" & Right("00" & Second(rstCM1("Ora_Fine_Pian")), 2)
   End If
    
    sist1 = Replace(rstCM1("Mezzo_trasmissione") & "", "'", "''")
    tel1 = Replace(rstCM1("Telefono_Mezzo_tx") & "", "'", "''")
    loc1 = Replace(rstCM1("Destinazione_collegamento") & "", "'", "''")
    txp1 = Replace(rstCM1("Transponder") & "", "'", "''")
    audio1_1 = Replace(rstCM1("Audio_Stereo1") & "", "'", "''")
    audio2_1 = Replace(rstCM1("Audio_Stereo2") & "", "'", "''")
    note1 = Replace(rstCM1("Note_output") & "", "'", "''")
    
    '3-3-2005 AGGIUNTI CAMPI NUOVI
      
    satellite1 = Replace(rstCM1("Satellite") & "", "'", "''")
    'VERIFICARE CON ANDREOLI
    NRO_Cod_satellite1 = Replace(rstCM1("Codif_Sat") & "", "'", "''")
    Cod_satellite1 = ""
    If Trim(NRO_Cod_satellite1) <> "" Then
        If InStr(NRO_Cod_satellite1, "-") > 0 Then
            d = Split(NRO_Cod_satellite1, "-")
            If LBound(d) = 0 Then
                Cod_satellite1 = d(0)
            End If
            If UBound(d) > 0 Then
                NRO_Cod_satellite1 = d(1)
            End If
        End If
    End If
    Posizione_orbitale1 = Replace(rstCM1("Posizione_orbitale") & "", "'", "''")
    Canale1 = Replace(rstCM1("Canale") & "", "'", "''")
    System1 = Replace(rstCM1("Sistema_tx_sat") & "", "'", "''")
    Frequenza1 = Replace(rstCM1("Frequenza") & "", "'", "''")
    Polarizzazione1 = Replace(rstCM1("Polarizzazione") & "", "'", "''")
    Bit_Rate1 = Replace(rstCM1("BitRate") & "", "'", "''")
    Symbol_Rate1 = Replace(rstCM1("SymbolRate") & "", "'", "''")
    MTT1 = Replace(rstCM1("Multiplex_Tx_Type") & "", "'", "''")
    FEC1 = Replace(rstCM1("FEC") & "", "'", "''")
    Origine1 = Replace(rstCM1("Origine_collegamento") & "", "'", "''")
    Destinazione1 = Replace(rstCM1("Destinazione_collegamento") & "", "'", "''")
    Miaf1 = Replace(rstCM1("Ind_Miaf") & "", "'", "''")
    Mezzo_ricez1 = Replace(rstCM1("Mezzo_ricezione") & "", "'", "''")
    Banda1 = Replace(rstCM1("Banda") & "", "'", "''")
    CM_Freq_Dw_link = Replace(rstCM1("frequenzadown") & "" & "", "'", "''")
    CM_Pol_Dw_link = Replace(rstCM1("polarizzazionedown") & "" & "", "'", "''")
    CM_Codice_Servizio = Replace(rstCM1("codiceservizio") & "" & "", "'", "''")
End If

If IsObject(rstCM2) Then
    If Not rstCM2.EOF Then
    'collegamenti mobili C2
    'C2_Ora_Iniz_PM = rstCM2("Ora_Inizio_Pian") & ""
    'C2_Ora_Fine_PM = rstCM2("Ora_Fine_Pian") & ""
    
   If IsNull(rstCM2("Ora_Inizio_Pian")) Then
        C2_Ora_Iniz_PM = "1900/01/01 00:00:00"
   Else
        C2_Ora_Iniz_PM = "1900/01/01 " & Right("00" & Hour(rstCM2("Ora_Inizio_Pian")), 2) & ":" & Right("00" & Minute(rstCM2("Ora_Inizio_Pian")), 2) & ":" & Right("00" & Second(rstCM2("Ora_Inizio_Pian")), 2)
   End If
        
    
   If IsNull(rstCM2("Ora_Fine_Pian")) Then
        C2_Ora_Fine_PM = "1900/01/01 00:00:00"
   Else
        C2_Ora_Fine_PM = "1900/01/01 " & Right("00" & Hour(rstCM2("Ora_Fine_Pian")), 2) & ":" & Right("00" & Minute(rstCM2("Ora_Fine_Pian")), 2) & ":" & Right("00" & Second(rstCM2("Ora_Fine_Pian")), 2)
   End If
    
    sist2 = Replace(rstCM2("Mezzo_trasmissione") & "", "'", "''")
    tel2 = Replace(rstCM2("Telefono_Mezzo_tx") & "", "'", "''")
    loc2 = Replace(rstCM2("Destinazione_collegamento") & "", "'", "''")
    txp2 = Replace(rstCM2("Transponder") & "", "'", "''")
    audio1_2 = Replace(rstCM2("Audio_Stereo1") & "", "'", "''")
    audio2_2 = Replace(rstCM2("Audio_Stereo2") & "", "'", "''")
    note2 = Replace(rstCM2("Note_output") & "", "'", "''")
    
    '3-3-2005 AGGIUNTI CAMPI NUOVI
    
    satellite2 = Replace(rstCM2("Satellite") & "", "'", "''")
    'VERIFICARE CON ANDREOLI
    NRO_Cod_satellite2 = Replace(rstCM2("Codif_Sat") & "", "'", "''")
    Cod_satellite2 = ""
    If Trim(NRO_Cod_satellite2) <> "" Then
        If InStr(NRO_Cod_satellite2, "-") > 0 Then
            d = Split(NRO_Cod_satellite2, "-")
            If LBound(d) = 0 Then
                Cod_satellite2 = d(0)
            End If
            If UBound(d) > 0 Then
                NRO_Cod_satellite2 = d(1)
            End If
        End If
    End If
    Posizione_orbitale2 = Replace(rstCM2("Posizione_orbitale") & "", "'", "''")
    Canale2 = Replace(rstCM2("Canale") & "", "'", "''")
    System2 = Replace(rstCM2("Sistema_tx_sat") & "", "'", "''")
    Frequenza2 = Replace(rstCM2("Frequenza") & "", "'", "''")
    Polarizzazione2 = Replace(rstCM2("Polarizzazione") & "", "'", "''")
    Bit_Rate2 = Replace(rstCM2("BitRate") & "", "'", "''")
    Symbol_Rate2 = Replace(rstCM2("SymbolRate") & "", "'", "''")
    MTT2 = Replace(rstCM2("Multiplex_Tx_Type") & "", "'", "''")
    FEC2 = Replace(rstCM2("FEC") & "", "'", "''")
    Origine2 = Replace(rstCM2("Origine_collegamento") & "", "'", "''")
    Destinazione2 = Replace(rstCM2("Destinazione_collegamento") & "", "'", "''")
    Miaf2 = Replace(rstCM2("Ind_Miaf") & "", "'", "''")
    Mezzo_ricez2 = Replace(rstCM2("Mezzo_ricezione") & "", "'", "''")
    Banda2 = Replace(rstCM2("Banda") & "", "'", "''")
    C2_Freq_Dw_link = Replace(rstCM2("frequenzadown") & "", "'", "''")
    C2_Pol_Dw_link = Replace(rstCM2("polarizzazionedown") & "", "'", "''")
    C2_Codice_Servizio = Replace(rstCM2("codiceservizio") & "", "'", "''")
   End If
End If
    
    sql = "INSERT INTO CollegamentiMobili (" & _
        "Anno_Rich_CM, Num_Rich_CM, Ver_Rich_CM, Versione_CM," & _
        "CM_satellite, CM_Cod_satellite, CM_NRO_Cod_satellite ,CM_Posizione_orbitale, " & _
        "CM_Canale, CM_System, CM_Frequenza, CM_Polarizzazione, " & _
        "CM_Bit_Rate, CM_Symbol_Rate, CM_MTT , CM_FEC, " & _
        "CM_Origine, CM_Destinazione, C1_Miaf_CM, C1_Mezzo_ricez_CM , CM_Banda , " & _
        "C2_satellite, C2_Cod_satellite, C2_NRO_Cod_satellite ,C2_Posizione_orbitale, " & _
        "C2_Canale, C2_System, C2_Frequenza, C2_Polarizzazione, " & _
        "C2_Bit_Rate, C2_Symbol_Rate, C2_MTT , C2_FEC, " & _
        "C2_Origine, C2_Destinazione, C2_Miaf_CM, C2_Mezzo_ricez_CM , C2_Banda , " & _
        "C1_Sistema_CM, C1_Telefono_CM, C1_Nome_TXP_CM, C1_Localita_RX_CM," & "C1_Audio1_CM, C1_Audio2_CM, C1_Ora_Inizio_CM, C1_Ora_Fine_CM, C1_Note_CM," & "C2_Sistema_CM, C2_Telefono_CM, C2_Nome_TXP_CM, C2_Localita_RX_CM," & "C2_Audio1_CM, C2_Audio2_CM, C2_Ora_Inizio_CM, C2_Ora_Fine_CM, C2_Note_CM, CM_Freq_Dw_link, CM_Pol_Dw_link, CM_Codice_Servizio, C2_Freq_Dw_link, C2_Pol_Dw_link, C2_Codice_Servizio" & _
        ") VALUES (" & _
            ANNO_RICHIESTA & ", " & NUM_RICHIESTA & ", " & ver_richiesta & ", " & ver & " , " & _
            "'" & satellite1 & "', '" & Cod_satellite1 & "', " & "'" & NRO_Cod_satellite1 & "', '" & Posizione_orbitale1 & "', " & _
            "'" & Canale1 & "', '" & System1 & "', " & "'" & Frequenza1 & "', '" & Polarizzazione1 & "', " & _
            "'" & Bit_Rate1 & "', '" & Symbol_Rate1 & "', " & "'" & MTT1 & "', '" & FEC1 & "', " & _
            "'" & Origine1 & "', '" & Destinazione1 & "', " & "'" & Miaf1 & "', '" & Mezzo_ricez1 & "', '" & Banda1 & "', " & _
            "'" & satellite2 & "', '" & Cod_satellite2 & "', " & "'" & NRO_Cod_satellite2 & "', '" & Posizione_orbitale2 & "', " & _
            "'" & Canale2 & "', '" & System2 & "', " & "'" & Frequenza2 & "', '" & Polarizzazione2 & "', " & _
            "'" & Bit_Rate2 & "', '" & Symbol_Rate2 & "', " & "'" & MTT2 & "', '" & FEC2 & "', " & _
            "'" & Origine2 & "', '" & Destinazione2 & "', " & "'" & Miaf2 & "', '" & Mezzo_ricez2 & "', '" & Banda2 & "', " & _
            "'" & sist1 & "', '" & tel1 & "', " & "'" & txp1 & "', '" & loc1 & "', " & "'" & audio1_1 & "', '" & audio2_1 & "', '" & CM_Ora_Iniz_PM & "', '" & CM_Ora_Fine_PM & "', " & "'" & note1 & "', " & "'" & sist2 & "', '" & tel2 & "', " & _
            "'" & txp2 & "', '" & loc2 & "', " & "'" & audio1_2 & "', '" & audio2_2 & "', '" & C2_Ora_Iniz_PM & "', '" & C2_Ora_Fine_PM & "', '" & note2 & "', '" & CM_Freq_Dw_link & "', '" & CM_Pol_Dw_link & "', '" & CM_Codice_Servizio & "', '" & C2_Freq_Dw_link & "', '" & C2_Pol_Dw_link & "', '" & C2_Codice_Servizio & "')"

    getInsertNewCollegamentoMobile = sql
    Exit Function
err_getInsertNewCollegamentoMobile:
    getInsertNewCollegamentoMobile = ""
    WriteLog "Errore nella funzione di preparazione stringa di inserimento CM (nuova) : " & Err.Description & " richiesta : " & ANNO_RICHIESTA & NUM_RICHIESTA
End Function

Function isStessiDati(rsRich As Object, rsRprec As Object) As Boolean
    isStessiDati = False
    On Error GoTo err_handler
    
    If rsRich("Data_Ora_Inizio") <> rsRprec("Data_Ora_Inizio") Then
        Exit Function
    End If
    If rsRich("Data_Ora_Fine") <> rsRprec("Data_Ora_Fine") Then
        Exit Function
    End If
    If rsRich("Comune_Orig") <> rsRprec("Comune_Orig") Then
        Exit Function
    End If
    If rsRich("Comune_Destin") <> rsRprec("Comune_Destin") Then
        Exit Function
    End If
    If rsRich("Cod_Ente") <> rsRprec("Cod_Ente") Then
        Exit Function
    End If
    If rsRich("Colleg") <> rsRprec("Colleg") Then
        Exit Function
    End If
    If rsRich("Orientamento") <> rsRprec("Orientamento") Then
        Exit Function
    End If
    If rsRich("Titolo_tx") <> rsRprec("Titolo_tx") Then
        Exit Function
    End If
    If rsRich("Cod_Tipo_tx") <> rsRprec("Cod_Tipo_tx") Then
        Exit Function
    End If
    If rsRich("Cod_Naz_Orig") <> rsRprec("Cod_Naz_Orig") Then
        Exit Function
    End If
    If rsRich("Cod_Localita_Orig") <> rsRprec("Cod_Localita_Orig") Then
        Exit Function
    End If
    If rsRich("Cod_Prov_Orig") <> rsRprec("Cod_Prov_Orig") Then
        Exit Function
    End If
    If rsRich("Indirizzo_Orig") <> rsRprec("Indirizzo_Orig") Then
        Exit Function
    End If
    If rsRich("Ente_Orig") <> rsRprec("Ente_Orig") Then
        Exit Function
    End If
    If rsRich("Standard") <> rsRprec("Standard") Then
        Exit Function
    End If
    If rsRich("Cod_Naz_Destin") <> rsRprec("Cod_Naz_Destin") Then
        Exit Function
    End If
    If rsRich("Cod_Localita_Destin") <> rsRprec("Cod_Localita_Destin") Then
        Exit Function
    End If
    If rsRich("Cod_Prov_Destin") <> rsRprec("Cod_Prov_Destin") Then
        Exit Function
    End If
    If rsRich("Indirizzo_Destin") <> rsRprec("Indirizzo_Destin") Then
        Exit Function
    End If
    If rsRich("Lavorazione") <> rsRprec("Lavorazione") Then
        Exit Function
    End If
    If rsRich("Indic_Cron_Origine") <> rsRprec("Indic_Cron_Origine") Then
        Exit Function
    End If
    If rsRich("Note") <> rsRprec("Note") Then
        Exit Function
    End If
    If rsRich("Data_Limite_Ripet") <> rsRprec("Data_Limite_Ripet") Then
        Exit Function
    End If
    If rsRich("Lunedi") <> rsRprec("Lunedi") Then
        Exit Function
    End If
    If rsRich("Martedi") <> rsRprec("Martedi") Then
        Exit Function
    End If
    If rsRich("Mercoledi") <> rsRprec("Mercoledi") Then
        Exit Function
    End If
    If rsRich("Giovedi") <> rsRprec("Giovedi") Then
        Exit Function
    End If
    If rsRich("Venerdi") <> rsRprec("Venerdi") Then
        Exit Function
    End If
    If rsRich("Sabato") <> rsRprec("Sabato") Then
        Exit Function
    End If
    If rsRich("Domenica") <> rsRprec("Domenica") Then
        Exit Function
    End If
    If rsRich("CM_Sistema") <> rsRprec("CM_Sistema") Then
        Exit Function
    End If
    If rsRich("CM_Num_Telefono") <> rsRprec("CM_Num_Telefono") Then
        Exit Function
    End If
    If rsRich("CM_Nome_TXP") <> rsRprec("CM_Nome_TXP") Then
        Exit Function
    End If
    If rsRich("CM_Localita_RX") <> rsRprec("CM_Localita_RX") Then
        Exit Function
    End If
    If rsRich("CM_Audio1") <> rsRprec("CM_Audio1") Then
        Exit Function
    End If
    If rsRich("CM_Audio2") <> rsRprec("CM_Audio2") Then
        Exit Function
    End If
    If rsRich("CM_Note") <> rsRprec("CM_Note") Then
        Exit Function
    End If
    If rsRich("CM_Ora_Inizio_PM") <> rsRprec("CM_Ora_Inizio_PM") Then
        Exit Function
    End If
    If rsRich("CM_Ora_Fine_PM") <> rsRprec("CM_Ora_Fine_PM") Then
        Exit Function
    End If
    isStessiDati = True
    Exit Function
err_handler:
    WriteLog "Errore nella funzione di confronto dei dati dei ponti : " & Err.Description
End Function

'// Non utilizzata
Function AbbinaOttimPrec(conn As Object, AnnR, NumR, VerR, msg As String) As Integer
Dim sqlCorr, sqlPrec, rsRisCorr, rsRichPrec As Object, rsRich As Object, sqlRich, rsRisPrec

On Error Resume Next
    AbbinaOttimPrec = 0

    sqlRich = "select * from Richiesta where  anno_rich = " & AnnR & " and num_rich = " & NumR & " and ver_rich = " & VerR
    Set rsRich = CreateObject("ADODB.Recordset")
    If Err <> 0 Then
        msg = "Errore creazione oggetto recordset richiesta."
        GoTo err_AbbinaOttimPrec
    End If
    rsRich.Open sqlRich, conn, 3
    If Err <> 0 Then
        msg = "Errore apertura recordset richiesta." & vbCrLf & sqlRich
        GoTo err_AbbinaOttimPrec
    End If
   
    If rsRich.EOF Then
        WriteLog "Non trovo la richiesta"
        rsRich.Close
        Set rsRich = Nothing
        Exit Function
    End If
      
    sqlCorr = "select * from Risorse where  anno_ottim is null and (indic_annullata is null) and id_anno_rich = " & AnnR & " and id_num_rich = " & NumR & " and id_ver_rich = " & VerR
    Set rsRisCorr = CreateObject("ADODB.Recordset")
    If Err <> 0 Then
       msg = "Errore creazione oggetto recordset risorse."
       GoTo err_AbbinaOttimPrec
    End If
    rsRisCorr.Open sqlCorr, conn, 3
    If Err <> 0 Then
       msg = "Errore apertura recordset risorse." & vbCrLf & sqlCorr
       GoTo err_AbbinaOttimPrec
    End If
    
    Dim vprec, cntSql, sqlExec(), sqlExecFono()
   
    cntSql = 0
    'cerca la eventuale versione precedente della richiesta
    vprec = getVersPrecRich(conn, AnnR, NumR, VerR)
    If vprec = 0 Then
        WriteLog "Versione precedente della richiesta non trovata"
        'versione precedente non trovata, non fa nulla
        GoTo exit_AbbinaOttimPrec
    End If
    'carica la versione precedente
    sqlPrec = "select * from Richiesta where  anno_rich = " & AnnR & " and num_rich = " & NumR & " and ver_rich = " & vprec
    Set rsRichPrec = CreateObject("ADODB.Recordset")
    If Err <> 0 Then
       msg = "Errore creazione oggetto recordset richiesta con versione precedente."
       GoTo err_AbbinaOttimPrec
    End If
    rsRichPrec.Open sqlPrec, conn, 3
    If Err <> 0 Then
       msg = "Errore apertura recordset richiesta con versione precedente." & vbCrLf & sqlPrec
       GoTo err_AbbinaOttimPrec
    End If
    
    If Not isStessiDati(rsRich, rsRichPrec) Then
    'origine  o destinazione o date diverse: dovrebbe dare un warning
       AbbinaOttimPrec = 2
       GoTo exit_AbbinaOttimPrec
    End If
    
    If rsRisCorr.EOF Then
        WriteLog "Non trovo le risorse della richiesta"
    End If

    'cicla sulle risorse della richiesta corrente
    Do While Not rsRisCorr.EOF
        msg = ""
      
        sqlPrec = "select * from Risorse where id_anno_rich = " & AnnR & _
        " and id_num_rich = " & NumR & " and id_ver_rich = " & vprec & _
        " and Cod_Risorsa = '" & rsRisCorr("Cod_Risorsa") & "'" & _
        " and Data_Ora_Inizio = '" & rsRisCorr("Data_Ora_Inizio") & "'" & _
        " and Data_Ora_Fine = '" & rsRisCorr("Data_Ora_Fine") & "'" & _
        " and quantita='" & rsRisCorr("quantita") & _
        "' and indic_annullata is null and anno_ottim is not null"
        Set rsRisPrec = CreateObject("ADODB.Recordset")
        If Err <> 0 Then
           msg = "Errore creazione oggetto recordset risorse precedenti.<br>" & Err.Description
           GoTo err_AbbinaOttimPrec
        End If
        rsRisPrec.Open sqlPrec, conn, 3
        If Err <> 0 Then
           msg = "Errore apertura recordset risorse precedenti.<br>" & Err.Description & "<br>" & sqlPrec
           GoTo err_AbbinaOttimPrec
        End If
        If Not rsRisPrec.EOF Then
        'memorizzo in array l'istruzione SQL per eseguire dopo l'update
            ReDim Preserve sqlExec(cntSql)
            ReDim Preserve sqlExecFono(cntSql)
            sqlExec(cntSql) = "Update risorse set " & _
            "Anno_Ottim=" & rsRisPrec("Anno_Ottim") & "," & _
            "Num_Ottim=" & rsRisPrec("Num_Ottim") & "," & _
            "Progressivo_Ottim=" & rsRisPrec("Progressivo_Ottim") & "," & _
            "Stato_Risorsa= '" & rsRisPrec("Stato_Risorsa") & "'" & _
            " Where progressivo=" & rsRisCorr("progressivo") & _
            " and anno_ottim is null and (indic_annullata is null)" & _
            " and id_anno_rich = " & AnnR & " and id_num_rich = " & NumR & " and id_ver_rich = " & VerR
            sqlExecFono(cntSql) = "Update Ottim_fono set " & _
            "flag_variante= ' ' where " & _
            "Anno_Ottim=" & rsRisPrec("Anno_Ottim") & " and " & _
            "Num_Ottim=" & rsRisPrec("Num_Ottim")
            cntSql = cntSql + 1
            If Err <> 0 Then
                msg = rsRisPrec("Anno_Ottim") & "/" & rsRisPrec("Num_Ottim") & "/" & rsRisPrec("progressivo_ottim") & "<br>Errore aggiornamento risorse da risorse precedenti.<br>" & Err.Description
                GoTo err_AbbinaOttimPrec
            End If
        End If
        rsRisPrec.Close
        Set rsRisPrec = Nothing
              
        rsRisCorr.movenext
    Loop
    rsRisCorr.Close
    Set rsRisCorr = Nothing
    Dim ik
    For ik = 0 To cntSql - 1
'esegue in loop gli update necessari
        conn.execute sqlExec(ik)
        If Err <> 0 Then
            msg = "Errore aggiornamento risorse da lettura risorse precedenti." & vbCrLf & sqlExec(ik)
            GoTo err_AbbinaOttimPrec
        End If
        conn.execute sqlExecFono(ik)
        If Err <> 0 Then
            msg = "Errore aggiornamento flag fono da lettura risorse precedenti." & vbCrLf & sqlExec(ik)
            GoTo err_AbbinaOttimPrec
        End If
    Next
    AbbinaOttimPrec = 1
   
exit_AbbinaOttimPrec:
    If IsObject(rsRisCorr) Then
        rsRisCorr.Close
        Set rsRisCorr = Nothing
    End If
    If IsObject(rsRisPrec) Then
        rsRisPrec.Close
        Set rsRisPrec = Nothing
    End If
    If IsObject(rsRichPrec) Then
        rsRichPrec.Close
        Set rsRichPrec = Nothing
    End If
    
    rsRich.Close
    Set rsRich = Nothing
    
    Exit Function
   
err_AbbinaOttimPrec:
    AbbinaOttimPrec = -1
    msg = msg & vbCrLf & Err.Description
    Resume exit_AbbinaOttimPrec
End Function


Private Function GetInsertStatementForMezzi_Ripresa(ByRef Richiesta As RaiBook.M7375350_RICHIESTE, Progressivo As Integer, ByVal vers As Integer) As String

   Dim sql As String

   With Richiesta
      sql = "INSERT INTO MezziRipresa VALUES(" & _
            .A7375351_ANNO_RICH & ", " & _
            .A7375351_NUM_RICH & ", " & _
            .A7375351_VER_RICH & ", " & _
            Progressivo + 1 & ", " & _
            "'" & Replace(.A7375351_EL_MR_MEZZO(Progressivo).A7375351_MR_MEZZO, "'", "''") & "', " & _
            "'" & Replace(.A7375351_EL_MR_MEZZO(Progressivo).A7375351_MR_DESCR, "'", "''") & "', " & _
            "'" & Replace(.A7375351_EL_MR_MEZZO(Progressivo).A7375351_MR_ALTRO, "'", "''") & "', " & _
            vers & ")"
   End With

   GetInsertStatementForMezzi_Ripresa = sql
End Function


Private Function convertiOre24(ByVal ora As String) As String
Dim hh As String
On Error GoTo err_handler

    hh = IIf(InStr(ora, ":") > 0, ora, Left(ora, 2) & ":" & Mid(ora, 3))
    convertiOre24 = IIf(hh = "24:00", "23:59:59", hh & ":00")
Exit Function

err_handler:
WriteLog " errore nella funzione di conversione delle ore " & Err.Description
End Function

Private Function leggiDataHost(ByVal data As String, ByVal ora As String) As String
Dim hh As String
   
      hh = Right(String(4, "0") & ora, 4)
      hh = Mid(hh, 1, 2) & ":" & Mid(hh, 3)
      hh = convertiOre24(hh)
      
      leggiDataHost = Mid(data, 1, 4) & "/" & Mid(data, 5, 2) & "/" & Mid(data, 7, 2) & " " & hh
End Function


Private Function GetInsertStatementForRichiesta(ByRef Richiesta As RaiBook.M7375350_RICHIESTE) As String
   
   Dim DataOraInizio As String
   Dim OraInizio As String
   Dim DataOraFine As String
   Dim OraFine As String
   Dim DataOraInoltro As String
   Dim OraInoltro As String
   Dim Data_Limite_Ripet As String
   Dim CM_Ora_Iniz_PM As String
   Dim CM_Ora_Fine_PM As String
   Dim C2_Ora_Iniz_PM As String
   Dim C2_Ora_Fine_PM As String
    
   Dim sql As String
   
   With Richiesta
   
      If .A7375351_ORA_FINE > 2400 Then
         .A7375351_ORA_FINE = 2400
      End If

   
      OraInizio = Right(String(4, "0") & .A7375351_ORA_INIZIO, 4)
      OraInizio = Mid(OraInizio, 1, 2) & ":" & Mid(OraInizio, 3)
      OraInizio = convertiOre24(OraInizio)
      
      OraFine = Right(String(4, "0") & .A7375351_ORA_FINE, 4)
      OraFine = Mid(OraFine, 1, 2) & ":" & Mid(OraFine, 3)
      OraFine = convertiOre24(OraFine)
        
      OraInoltro = Right(String(4, "0") & .A7375351_ORA_INOLTRO, 4)
      OraInoltro = Mid(OraInoltro, 1, 2) & ":" & Mid(OraInoltro, 3)
      OraInoltro = convertiOre24(OraInoltro)
      
      DataOraInizio = .A7375351_DATA_INIZIO & " " & OraInizio
      DataOraFine = .A7375351_DATA_FINE & " " & OraFine
      
'controlli sulle date non valorizzate
      If .A7375351_DATA_INOLTRO = 0 Then
        DataOraInoltro = "NULL"
      Else
        DataOraInoltro = "'" & .A7375351_DATA_INOLTRO & " " & OraInoltro & "'"
      End If
      
      If .A7375351_DATA_LIMITE_RIPET = 0 Then
         Data_Limite_Ripet = "NULL"
      Else
         Data_Limite_Ripet = "'" & .A7375351_DATA_LIMITE_RIPET & "'"
      End If
      
'collegamenti mobili 1
      If .A7375351_CM_ORA_INIZ_PM = 0 Then
         CM_Ora_Iniz_PM = "NULL"
      Else
         CM_Ora_Iniz_PM = Right(String(4, "0") & .A7375351_CM_ORA_INIZ_PM, 4)
         CM_Ora_Iniz_PM = "'" & convertiOre24(Mid(CM_Ora_Iniz_PM, 1, 2) & ":" & Mid(CM_Ora_Iniz_PM, 3)) & "'"
      End If
      
      If .A7375351_CM_ORA_FINE_PM = 0 Then
         CM_Ora_Fine_PM = "NULL"
      Else
         CM_Ora_Fine_PM = Right(String(4, "0") & .A7375351_CM_ORA_FINE_PM, 4)
         CM_Ora_Fine_PM = "'" & convertiOre24(Mid(CM_Ora_Fine_PM, 1, 2) & ":" & Mid(CM_Ora_Fine_PM, 3)) & "'"
      End If
      
'collegamenti mobili 2
      If .A7375351_C2_ORA_INIZ_PM = 0 Then
         C2_Ora_Iniz_PM = "'1900/01/01 00:00.0'"
      Else
         C2_Ora_Iniz_PM = Right(String(4, "0") & .A7375351_C2_ORA_INIZ_PM, 4)
         C2_Ora_Iniz_PM = "'" & convertiOre24(Mid(C2_Ora_Iniz_PM, 1, 2) & ":" & Mid(C2_Ora_Iniz_PM, 3)) & "'"
      End If
      
      If .A7375351_C2_ORA_FINE_PM = 0 Then
         C2_Ora_Fine_PM = "'1900/01/01 00:00.0'"
      Else
         C2_Ora_Fine_PM = Right(String(4, "0") & .A7375351_C2_ORA_FINE_PM, 4)
         C2_Ora_Fine_PM = "'" & convertiOre24(Mid(C2_Ora_Fine_PM, 1, 2) & ":" & Mid(C2_Ora_Fine_PM, 3)) & "'"
      End If
      
      sql = "INSERT INTO Richiesta VALUES(" & _
            .A7375351_ANNO_RICH & ", " & _
            .A7375351_NUM_RICH & ", " & _
            .A7375351_VER_RICH & ", " & _
            "'" & .A7375351_COD_TIPO_RICHIESTA & "', " & _
            .A7375351_VERSIONE_RIF & ", " & _
            "'" & .A7375351_COD_ENTE & "', " & _
            "'" & DataOraInizio & "', " & _
            "'" & DataOraFine & "', " & _
            "'" & .A7375351_COLLEG & "', " & _
            "'" & .A7375351_ORIENTAMENTO & "', " & _
            "'" & Replace(.A7375351_TITOLO_TX, "'", "''") & "', " & _
            "'" & .A7375351_COD_TIPO_TX & "', " & _
            "'" & .A7375351_COD_NAZ_ORIG & "', " & _
            "'" & .A7375351_COD_PROV_ORIG & "', " & _
            "'" & Replace(.A7375351_COMUNE_ORIG, "'", "''") & "', " & _
            "'" & Replace(.A7375351_INDIRIZZO_ORIG, "'", "''") & "', " & _
            "'" & .A7375351_ENTE_ORIG & "', " & _
            "'" & .A7375351_STANDARD & "', " & _
            "'" & .A7375351_COD_NAZ_DESTIN & "', " & _
            "'" & .A7375351_COD_PROV_DESTIN & "', " & _
            "'" & Replace(.A7375351_COMUNE_DESTIN, "'", "''") & "', " & _
            "'" & Replace(.A7375351_INDIRIZZO_DESTIN, "'", "''") & "', " & _
            "'" & .A7375351_LAVORAZIONE & "', " & _
            "'" & .A7375351_INDIC_CRON_ORIGINE & "', "
            
      sql = sql & _
            "'" & .A7375351_INDIC_CRON_RISERVA & "', " & _
            "'" & .A7375351_LOCAL_CRON_RISERVA & "', " & _
            "'" & .A7375351_INDIC_CRON_OFFTUBE & "', " & _
            "'" & .A7375351_LOCAL_CRON_OFFTUBE & "', " & _
            "'" & Replace(.A7375351_NOTE, "'", "''") & "', " & _
            "'" & Replace(.A7375351_NOME_RICHIED, "'", "''") & "', " & _
            "'" & .A7375351_TEL_RICHIED & "', " & _
            Data_Limite_Ripet & ", " & _
            "'" & .A7375351_LUNEDI & "', " & _
            "'" & .A7375351_MARTEDI & "', " & _
            "'" & .A7375351_MERCOLEDI & "', " & _
            "'" & .A7375351_GIOVEDI & "', " & _
            "'" & .A7375351_VENERDI & "', " & _
            "'" & .A7375351_SABATO & "', " & _
            "'" & .A7375351_DOMENICA & "', " & _
            "'" & .A7375351_COD_TIPO_SEGNALAZ_C & "', " & _
            .A7375351_VERSIONE_RIF_C & ", " & _
            "'" & .A7375351_COD_TIPO_SEGNALAZ_F & "', " & _
            .A7375351_VERSIONE_RIF_F & ", " & _
            "'" & .A7375351_COD_TIPO_SEGNALAZ_M & "', " & _
            .A7375351_VERSIONE_RIF_M & ", " & _
            "'" & .A7375351_COD_TIPO_SEGNALAZ_P & "', " & _
            .A7375351_VERSIONE_RIF_P & ", " & _
            "'" & .A7375351_STATO_CIRCUITI & "', "
            
      sql = sql & _
            "'" & .A7375351_STATO_FACILITIES & "', " & _
            "'" & .A7375351_STATO_MEZZI & "', " & _
            "'" & .A7375351_STATO_PONTI & "', " & _
            "'" & Replace(.A7375351_CM_SISTEMA, "'", "''") & "', " & _
            "'" & Replace(.A7375351_CM_NUM_TELEFONO, "'", "''") & "', " & _
            "'" & Replace(.A7375351_CM_NOME_TXP, "'", "''") & "', " & _
            "'" & Replace(.A7375351_CM_LOCALITA_RX, "'", "''") & "', " & _
            "'" & .A7375351_CM_AUDIO1 & "', " & _
            "'" & .A7375351_CM_AUDIO2 & "', " & _
            CM_Ora_Iniz_PM & ", " & _
            CM_Ora_Fine_PM & ", '" & Replace(.A7375351_CM_NOTE & "", "'", "''") & "'," & _
            "'N', 'N', " & _
            "'" & Format(Now, "yyyymmdd hh:mm:ss") & "', " & _
            "NULL, " & _
            "'" & .A7375351_COD_LOC_ORIG & "', " & _
            "'" & .A7375351_COD_LOC_DESTIN & "', " & _
            "'" & .A7375351_FLAG_NOVITA & "', " & _
            "'" & .A7375351_UORG & "', " & _
            "'" & .A7375351_MATRICOLA & "', " & _
            "'" & .A7375351_SERIE & "', " & _
            "'" & .A7375351_PUNTATA & "', " & _
            "'" & .A7375351_DF & "'," & DataOraInoltro & ")"
   End With
   ' aggiungo dopo COD_LOC_DESTIN modifica per FLAG_NOVITA e centro di costo 18-6-02
   ' aggiungo vicino allo stato_web lo stato_monitor=N
   GetInsertStatementForRichiesta = sql

End Function

Private Function GetUpdateSuperateForRichiesta(ByRef Richiesta As RaiBook.M7375350_RICHIESTE) As String
    With Richiesta
    ' giulia  5-12-2002
    ' update della richiesta superata (aggiorna il campo cod_tipo_richiesta a superata)
        Dim Sql_update_sup As String
        Sql_update_sup = "UPDATE Richiesta SET " & _
                         "Cod_Tipo_Richiesta = 'S' " & _
                         "WHERE Anno_Rich = " & .A7375351_ANNO_RICH & " AND " & _
                         "Num_Rich = " & .A7375351_NUM_RICH & " AND " & _
                         "Ver_Rich <= " & .A7375351_VERSIONE_RIF
        'WriteLog "stringa update richiesta: " & Sql_update_sup
    End With
    GetUpdateSuperateForRichiesta = Sql_update_sup
End Function

Private Function GetUpdateForRisorseSuperate(ByRef Richiesta As RaiBook.M7375350_RICHIESTE) As String
    Dim Sql_update_sup As String
    With Richiesta
    ' G & R  15-9-2003
    ' update delle risorse di una richiesta superata (
    ' aggiorna il campo Stato_Risorsa a superata)
    
    Sql_update_sup = "UPDATE Risorse SET Stato_Risorsa = 'S' " & _
          "WHERE ID_Anno_Rich = " & .A7375351_ANNO_RICH & " AND " & _
          "ID_Num_Rich = " & .A7375351_NUM_RICH & " AND " & _
          "ID_Ver_Rich <= " & .A7375351_VERSIONE_RIF & ""
       
    End With
    GetUpdateForRisorseSuperate = Sql_update_sup
End Function


Private Function GetUpdateStatementForRichiesta(ByRef Richiesta As RaiBook.M7375350_RICHIESTE) As String
   
    Dim DataOraInizio As String
    Dim OraInizio As String
    Dim DataOraFine As String
    Dim OraFine As String
    Dim Data_Limite_Ripet As String
    Dim CM_Ora_Iniz_PM As String
    Dim CM_Ora_Fine_PM As String
    Dim DataOraInoltro As String
    Dim OraInoltro As String
     
    Dim sql As String
    
    With Richiesta
   
        OraInizio = Right(String(4, "0") & .A7375351_ORA_INIZIO, 4)
        OraInizio = Mid(OraInizio, 1, 2) & ":" & Mid(OraInizio, 3)
        OraInizio = convertiOre24(OraInizio)
        
        OraFine = Right(String(4, "0") & .A7375351_ORA_FINE, 4)
        OraFine = Mid(OraFine, 1, 2) & ":" & Mid(OraFine, 3)
        OraFine = convertiOre24(OraFine)
        
        OraInoltro = Right(String(4, "0") & .A7375351_ORA_INOLTRO, 4)
        OraInoltro = Mid(OraInoltro, 1, 2) & ":" & Mid(OraInoltro, 3)
        OraInoltro = convertiOre24(OraInoltro)
        
        DataOraInizio = .A7375351_DATA_INIZIO & " " & OraInizio
        DataOraFine = .A7375351_DATA_FINE & " " & OraFine

'controlli sulle date non valorizzate
      If .A7375351_DATA_INOLTRO = 0 Then
        DataOraInoltro = "NULL"
      Else
        DataOraInoltro = "'" & .A7375351_DATA_INOLTRO & " " & OraInoltro & "'"
      End If
        
        If .A7375351_DATA_LIMITE_RIPET = 0 Then
           Data_Limite_Ripet = "NULL"
        Else
           Data_Limite_Ripet = "'" & .A7375351_DATA_LIMITE_RIPET & "'"
        End If
        
        If .A7375351_CM_ORA_INIZ_PM = 0 Then
           CM_Ora_Iniz_PM = "NULL"
        Else
          CM_Ora_Iniz_PM = Right(String(4, "0") & .A7375351_CM_ORA_INIZ_PM, 4)
          CM_Ora_Iniz_PM = "'" & convertiOre24(Mid(CM_Ora_Iniz_PM, 1, 2) & ":" & Mid(CM_Ora_Iniz_PM, 3)) & "'"
        End If
      
        If .A7375351_CM_ORA_FINE_PM = 0 Then
            CM_Ora_Fine_PM = "NULL"
        Else
            CM_Ora_Fine_PM = Right(String(4, "0") & .A7375351_CM_ORA_FINE_PM, 4)
            CM_Ora_Fine_PM = "'" & convertiOre24(Mid(CM_Ora_Fine_PM, 1, 2) & ":" & Mid(CM_Ora_Fine_PM, 3)) & "'"
        End If
      
        sql = "UPDATE Richiesta SET " & _
            "Cod_Tipo_Richiesta = '" & .A7375351_COD_TIPO_RICHIESTA & "', " & _
            "Versione_Rif = " & .A7375351_VERSIONE_RIF & ", " & _
            "Cod_Ente = '" & .A7375351_COD_ENTE & "', " & _
            "Data_Ora_Inizio = '" & DataOraInizio & "', " & _
            "Data_Ora_Fine = '" & DataOraFine & "', " & _
            "Data_Ora_Inoltro = " & DataOraInoltro & ", " & _
            "Colleg = '" & .A7375351_COLLEG & "', " & _
            "Orientamento = '" & .A7375351_ORIENTAMENTO & "', " & _
            "Titolo_TX = '" & Replace(.A7375351_TITOLO_TX, "'", "''") & "', " & _
            "Cod_Tipo_TX = '" & .A7375351_COD_TIPO_TX & "', " & _
            "Cod_Naz_Orig = '" & .A7375351_COD_NAZ_ORIG & "', " & _
            "Cod_Prov_Orig = '" & .A7375351_COD_PROV_ORIG & "', " & _
            "Comune_Orig = '" & Replace(.A7375351_COMUNE_ORIG, "'", "''") & "', " & _
            "Indirizzo_Orig = '" & Replace(.A7375351_INDIRIZZO_ORIG, "'", "''") & "', " & _
            "Ente_Orig = '" & .A7375351_ENTE_ORIG & "', " & _
            "Standard = '" & .A7375351_STANDARD & "', " & _
            "Cod_Naz_Destin = '" & .A7375351_COD_NAZ_DESTIN & "', " & _
            "Cod_Prov_Destin = '" & .A7375351_COD_PROV_DESTIN & "', " & _
            "Comune_Destin = '" & Replace(.A7375351_COMUNE_DESTIN, "'", "''") & "', " & _
            "Indirizzo_Destin = '" & Replace(.A7375351_INDIRIZZO_DESTIN, "'", "''") & "', " & _
            "Lavorazione = '" & .A7375351_LAVORAZIONE & "', " & _
            "CM_Note = '" & Replace(.A7375351_CM_NOTE & "", "'", "''") & "', " & _
            "Indic_Cron_Origine = '" & .A7375351_INDIC_CRON_ORIGINE & "', "
            
      sql = sql & _
            "Indic_Cron_Riserva = '" & .A7375351_INDIC_CRON_RISERVA & "', " & _
            "Local_Cron_Riserva = '" & .A7375351_LOCAL_CRON_RISERVA & "', " & _
            "Indic_Cron_Offtube = '" & .A7375351_INDIC_CRON_OFFTUBE & "', " & _
            "Local_Cron_Offtube = '" & .A7375351_LOCAL_CRON_OFFTUBE & "', " & _
            "Note = '" & Replace(.A7375351_NOTE, "'", "''") & "', " & _
            "Nome_Richied = '" & Replace(.A7375351_NOME_RICHIED, "'", "''") & "', " & _
            "Tel_Richied = '" & .A7375351_TEL_RICHIED & "', " & _
            "Data_Limite_Ripet = " & Data_Limite_Ripet & ", " & _
            "Lunedi = '" & .A7375351_LUNEDI & "', " & _
            "Martedi = '" & .A7375351_MARTEDI & "', " & _
            "Mercoledi = '" & .A7375351_MERCOLEDI & "', " & _
            "Giovedi = '" & .A7375351_GIOVEDI & "', " & _
            "Venerdi = '" & .A7375351_VENERDI & "', " & _
            "Sabato = '" & .A7375351_SABATO & "', " & _
            "Domenica = '" & .A7375351_DOMENICA & "', " & _
            "Cod_Tipo_Segnalaz_C = '" & .A7375351_COD_TIPO_SEGNALAZ_C & "', " & _
            "Versione_Rif_C = " & .A7375351_VERSIONE_RIF_C & ", " & _
            "Cod_Tipo_Segnalaz_F = '" & .A7375351_COD_TIPO_SEGNALAZ_F & "', " & _
            "Versione_Rif_F = " & .A7375351_VERSIONE_RIF_F & ", " & _
            "Cod_Tipo_Segnalaz_M = '" & .A7375351_COD_TIPO_SEGNALAZ_M & "', " & _
            "Versione_Rif_M = " & .A7375351_VERSIONE_RIF_M & ", " & _
            "Cod_Tipo_Segnalaz_P = '" & .A7375351_COD_TIPO_SEGNALAZ_P & "', " & _
            "Versione_Rif_P = " & .A7375351_VERSIONE_RIF_P & ", " & _
            "Stato_Circuiti = '" & .A7375351_STATO_CIRCUITI & "', "
      sql = sql & _
            "Stato_Facilities = '" & .A7375351_STATO_FACILITIES & "', " & _
            "Stato_Mezzi = '" & .A7375351_STATO_MEZZI & "', " & _
            "Stato_Ponti = '" & .A7375351_STATO_PONTI & "', " & _
            "CM_Sistema = '" & Replace(.A7375351_CM_SISTEMA, "'", "''") & "', " & _
            "CM_Num_Telefono = '" & Replace(.A7375351_CM_NUM_TELEFONO, "'", "''") & "', " & _
            "CM_Nome_TXP = '" & Replace(.A7375351_CM_NOME_TXP, "'", "''") & "', " & _
            "CM_Localita_RX = '" & Replace(.A7375351_CM_LOCALITA_RX, "'", "''") & "', " & _
            "CM_Audio1 = '" & .A7375351_CM_AUDIO1 & "', " & _
            "CM_Audio2 = '" & .A7375351_CM_AUDIO2 & "', " & _
            "CM_Ora_Inizio_PM = " & CM_Ora_Iniz_PM & ", " & _
            "CM_Ora_Fine_PM = " & CM_Ora_Fine_PM & ", " & _
            "Stato_Monitor = 'N', " & _
            "Data_Ora_Acquisizione = '" & Format(Now, "yyyymmdd hh:mm:ss") & "', " & _
            "Id_Utente = 99, " & _
            "Cod_localita_orig = '" & .A7375351_COD_LOC_ORIG & "', " & _
            "Cod_localita_destin = '" & .A7375351_COD_LOC_DESTIN & "', "
 '---> tolto xche` in UPDATE non deve azzerare l'ultimo id_utente
 '           "Id_Utente = NULL, " & _
 '--- 1/9/2003 sostituito con 99 che e` l'id del motore <---

   ' aggiungo dopo COD_LOC_DESTIN modifica per FLAG_NOVITA e centro di costo 18-6-02
   ' aggiungo la gestione di stato_monitor, e da valutare se lo stato_web diversificato
     sql = sql & _
            "Flag_Novita = '" & .A7375351_FLAG_NOVITA & "', " & _
            "UORG = '" & .A7375351_UORG & "', " & _
            "Matricola = '" & .A7375351_MATRICOLA & "', " & _
            "Serie = '" & .A7375351_SERIE & "', " & _
            "Puntata = '" & .A7375351_PUNTATA & "', " & _
            "DF = '" & .A7375351_DF & "' " & _
            "WHERE Anno_Rich = " & .A7375351_ANNO_RICH & " AND " & _
            "Num_Rich = " & .A7375351_NUM_RICH & " AND " & _
            "Ver_Rich = " & .A7375351_VER_RICH

   End With
   GetUpdateStatementForRichiesta = sql
End Function


Private Function GetInsertStatementForRisorse(ByRef Richiesta As RaiBook.M7375350_RICHIESTE, Progressivo As Integer, ProgCommarea As Integer) As String
   
   Dim DataOraInizio As String
   Dim DataOraFine As String
   Dim OraInizio As String
   Dim OraFine As String
   
   Dim sql As String

   With Richiesta
   
      OraInizio = Right(String(4, "0") & .A7375351_EL_TAB_RIS(ProgCommarea).A7375351_EL_ORA_INIZ, 4)
      OraInizio = convertiOre24(OraInizio)
      DataOraInizio = .A7375351_EL_TAB_RIS(ProgCommarea).A7375351_EL_DATA_INIZ & " " & OraInizio
   
      OraFine = Right(String(4, "0") & .A7375351_EL_TAB_RIS(ProgCommarea).A7375351_EL_ORA_FINE, 4)
      OraFine = convertiOre24(OraFine)
      DataOraFine = .A7375351_EL_TAB_RIS(ProgCommarea).A7375351_EL_DATA_FINE & " " & OraFine
   
      sql = "INSERT INTO Risorse (ID_ANNO_RICH, ID_NUM_RICH, ID_VER_RICH, " & _
            "PROGRESSIVO, COD_RISORSA, QUANTITA, DATA_ORA_INIZIO, DATA_ORA_FINE, " & _
            "STATO_RISORSA) VALUES(" & _
            .A7375351_ANNO_RICH & ", " & _
            .A7375351_NUM_RICH & ", " & _
            .A7375351_VER_RICH & ", " & _
            Progressivo + 1 & ", " & _
            "'" & Replace(.A7375351_EL_TAB_RIS(ProgCommarea).A7375351_EL_COD, "'", "''") & "', " & _
            "'" & .A7375351_EL_TAB_RIS(ProgCommarea).A7375351_EL_QTA & "', " & _
            "'" & DataOraInizio & "', " & _
            "'" & DataOraFine & "', " & _
            "'N')"
            '''' "'N', NULL, NULL, NULL, NULL, NULL, NULL, NULL,NULL,NULL,'N')"
   End With

   GetInsertStatementForRisorse = sql
End Function

Private Function GetUpdateStatementForRisorse(ByRef Richiesta As RaiBook.M7375350_RICHIESTE, Progressivo As Integer) As String
   
   Dim DataOraInizio As String
   Dim DataOraFine As String
   Dim OraInizio As String
   Dim OraFine As String
   
   Dim sql As String

   With Richiesta
   
      OraInizio = Right(String(4, "0") & .A7375351_EL_TAB_RIS(Progressivo).A7375351_EL_ORA_INIZ, 4)
      OraInizio = convertiOre24(OraInizio)
      DataOraInizio = .A7375351_EL_TAB_RIS(Progressivo).A7375351_EL_DATA_INIZ & " " & OraInizio
   
      OraFine = Right(String(4, "0") & .A7375351_EL_TAB_RIS(Progressivo).A7375351_EL_ORA_FINE, 4)
      OraFine = convertiOre24(OraFine)
      DataOraFine = .A7375351_EL_TAB_RIS(Progressivo).A7375351_EL_DATA_FINE & " " & OraFine
   
      sql = "UPDATE Risorse SET " & _
            "Cod_risorsa = '" & Replace(.A7375351_EL_TAB_RIS(Progressivo).A7375351_EL_COD, "'", "''") & "', " & _
            "Quantita = '" & .A7375351_EL_TAB_RIS(Progressivo).A7375351_EL_QTA & "', " & _
            "Data_Ora_Inizio = '" & DataOraInizio & "', " & _
            "Data_Ora_Fine = '" & DataOraFine & "', " & _
            "Stato_risorsa = 'N', " & _
            "Anno_Ottim = NULL, " & _
            "Num_Ottim = NULL " & _
            "WHERE ID_Anno_Rich = " & .A7375351_ANNO_RICH & " AND " & _
            "ID_Num_Rich = " & .A7375351_NUM_RICH & " AND " & _
            "ID_Ver_Rich = " & .A7375351_VER_RICH & " AND " & _
            "Progressivo = " & Progressivo + 1
   End With

   GetUpdateStatementForRisorse = sql
End Function


Private Function GetLogFileName() As String

   Const INI_LOG_SECTION As String = "log"
   Const LOG_FILE = "log_file"
   Const CT_BEGIN_TAG As String = "$("
   Const CT_END_TAG As String = ")"
   Const CT_APP_PATH As String = "AppPath"

   Static AppLogFile As String
   '''Dim AppIniFile As String
   Dim CustomTag As String
   
   Dim s As String
   
   Dim N As Integer
   Dim i As Integer
   Dim ret As Long
   
   If AppLogFile = "" Then
      '''AppIniFile = GetIniFileName
      
      AppLogFile = Space(255)
      N = Len(AppLogFile)
      ret = GetPrivateProfileString(INI_LOG_SECTION, LOG_FILE, "", AppLogFile, N, AppIniFile)
      AppLogFile = Mid(AppLogFile, 1, ret)
      
      N = InStr(AppLogFile, CT_BEGIN_TAG)
      If N <> 0 Then
         i = InStr(N, AppLogFile, CT_END_TAG)
         N = N + Len(CT_BEGIN_TAG)
         CustomTag = Mid(AppLogFile, N, i - N)
         Select Case UCase(CustomTag)
            Case UCase(CT_APP_PATH)
               s = App.Path
               If Right(s, 1) = "\" Then s = Mid(s, 1, Len(s) - 1)
               AppLogFile = s & Mid(AppLogFile, i + 1)
         End Select
      End If
      
   End If
   GetLogFileName = AppLogFile
   
End Function

Function SetStatoMonitor(conn As Object, esito As String, descrizione As String) As Boolean

SetStatoMonitor = False

   Const MOTORE As String = "C"

   Dim sql As String
   Dim RecordsAffected As Long
   Dim strerror As String
   
   If descrizione <> "" Then
   Else
      descrizione = Space(50)
   End If
   sql = "UPDATE StatoMonitor SET " & _
                "Data_Ora = '" & CreaDataPerSQL(Now) & "', " & _
                "Esito = '" & esito & "', " & _
                "Descrizione = '" & Replace(Left(descrizione, 50), "'", "''") & "' " & _
         "WHERE ID_Motore='" & MOTORE & "'"
         'WriteLog sql
On Error GoTo err_handler
conn.execute sql, RecordsAffected
'
   If RecordsAffected <> 1 Then
      sql = "INSERT INTO StatoMonitor VALUES (" & _
            "'" & MOTORE & "', " & _
            "'" & CreaDataPerSQL(Now) & "', " & _
            "'" & esito & "', " & _
            "'" & Replace(Left(descrizione, 50), "'", "''") & "')"
      'WriteLog sql
      conn.execute sql
   End If

SetStatoMonitor = True
    Exit Function
    
err_handler:

    WriteLog "Errore nella funzione di aggiornamento dello stato del Monitor"
    If Err.Number <> 0 Or conn.errors.Count > 0 Then
        strerror = "Error #" & Err.Number & vbCrLf & _
              "   (Description: " & Err.Description & ")" & vbCrLf & _
              "   (Source: " & Err.Source & ")" & vbCrLf
        strerror = strerror & "    (Error count : " & conn.errors.Count & ")" & vbCrLf
        strerror = strerror & "    (NativeError: " & conn.errors.Item(0).NativeError & ")" & vbCrLf
        strerror = strerror & "    (NE Description: " & conn.errors.Item(0).Description & ")" & vbCrLf
        strerror = strerror & "    (SQLState: " & conn.errors.Item(0).SQLState & ")" & vbCrLf
        WriteLog (strerror)
        WriteLog vbCrLf & "Comando SQL: " & sql
    End If

End Function
Private Function DeleteStatoMonitor(conn As Object) As Boolean
    'cancella record elaborazione //
    DeleteStatoMonitor = False
    On Error GoTo err_handler
    
    Dim sql As String
    Dim strerror As String
    Dim rsMonitor As Object
    
    
    sql = "Delete from StatoMonitor where ID_MOTORE = 'T'"
    conn.execute sql
    
    DeleteStatoMonitor = True
    Exit Function
    
err_handler:

    WriteLog "Errore nella funzione di Cancellazione del record di elaborazione"
    If Err.Number <> 0 Or conn.errors.Count > 0 Then
        strerror = "Error #" & Err.Number & vbCrLf & _
              "   (Description: " & Err.Description & ")" & vbCrLf & _
              "   (Source: " & Err.Source & ")" & vbCrLf
        strerror = strerror & "    (Error count : " & conn.errors.Count & ")" & vbCrLf
        strerror = strerror & "    (NativeError: " & conn.errors.Item(0).NativeError & ")" & vbCrLf
        strerror = strerror & "    (NE Description: " & conn.errors.Item(0).Description & ")" & vbCrLf
        strerror = strerror & "    (SQLState: " & conn.errors.Item(0).SQLState & ")" & vbCrLf
        WriteLog (strerror)
        WriteLog vbCrLf & "Comando SQL: " & sql
    End If
End Function
Private Function DeleteOccupate(conn As Object) As Boolean
    'cancella eventuali occupazioni //
    DeleteOccupate = False
    On Error GoTo err_handler
    
    Dim sql As String
    Dim strerror As String
    
    sql = "Delete from RichiesteOccupate where Id = 99"
    conn.execute sql
    
    DeleteOccupate = True
Exit Function
    
err_handler:
    If Err.Number <> 0 Or conn.errors.Count > 0 Then
        strerror = "Error #" & Err.Number & vbCrLf & _
              "   (Description: " & Err.Description & ")" & vbCrLf & _
              "   (Source: " & Err.Source & ")" & vbCrLf
        strerror = strerror & "    (Error count : " & conn.errors.Count & ")" & vbCrLf
        strerror = strerror & "    (NativeError: " & conn.errors.Item(0).NativeError & ")" & vbCrLf
        strerror = strerror & "    (NE Description: " & conn.errors.Item(0).Description & ")" & vbCrLf
        strerror = strerror & "    (SQLState: " & conn.errors.Item(0).SQLState & ")" & vbCrLf
        WriteLog (strerror)
        WriteLog vbCrLf & "Comando SQL: " & sql
        Exit Function
    End If
End Function
Private Function funct_rollback(conn As Object) As Boolean
Dim strerror As String
funct_rollback = False

On Error GoTo err_handler
    
    conn.RollbackTrans
    funct_rollback = True
    Exit Function
    
err_handler:
WriteLog "errore in fase di esecuzione della rollback: verificare la richiesta corrente!"
    If Err.Number <> 0 Or conn.errors.Count > 0 Then
        strerror = "Error #" & Err.Number & vbCrLf & _
              "   (Description: " & Err.Description & ")" & vbCrLf & _
              "   (Source: " & Err.Source & ")" & vbCrLf
        strerror = strerror & "    (Error count : " & conn.errors.Count & ")" & vbCrLf
        strerror = strerror & "    (NativeError: " & conn.errors.Item(0).NativeError & ")" & vbCrLf
        strerror = strerror & "    (NE Description: " & conn.errors.Item(0).Description & ")" & vbCrLf
        strerror = strerror & "    (SQLState: " & conn.errors.Item(0).SQLState & ")" & vbCrLf
        WriteLog (strerror)
    End If
End Function

Private Function CtrlStatoMonitor(conn As Object) As Boolean
    'controllo record elaborazione // inserito per gestire la concorrenza con gli altri motori
    'se e' presente il record un'altra istanza del motore sta ancora girando
    
    CtrlStatoMonitor = False
    On Error GoTo err_handler
    
    Dim sql As String
    Dim RecordsAffected As Long
    Dim strerror As String
    Dim rsMonitor As Object
    Dim data_corr As String
    
    
    sql = "Select * from StatoMonitor where ID_MOTORE = 'T'"
    Set rsMonitor = conn.execute(sql)
    If rsMonitor.EOF Then
    Else
        'fare una verifica:  se il tempo di elaborazione supera 5 minuti cancellare il record
        data_corr = CreaDataPerSQL(Now)
        If DateDiff("n", rsMonitor("data_ora"), data_corr) > 5 Then
            WriteLog "L'elaborazione e' rimasta appesa, il record di elaborazione viene eliminato"
            conn.execute "Delete StatoMonitor where ID_Motore = 'T'"
        Else
            Exit Function 'c'e' un'altra elaborazione in corso
        End If
    End If
    
    CtrlStatoMonitor = True
    Exit Function
    
err_handler:
    If Err.Number <> 0 Or conn.errors.Count > 0 Then
        strerror = "Error #" & Err.Number & vbCrLf & _
              "   (Description: " & Err.Description & ")" & vbCrLf & _
              "   (Source: " & Err.Source & ")" & vbCrLf
        strerror = strerror & "    (Error count : " & conn.errors.Count & ")" & vbCrLf
        strerror = strerror & "    (NativeError: " & conn.errors.Item(0).NativeError & ")" & vbCrLf
        strerror = strerror & "    (NE Description: " & conn.errors.Item(0).Description & ")" & vbCrLf
        strerror = strerror & "    (SQLState: " & conn.errors.Item(0).SQLState & ")" & vbCrLf
        WriteLog (strerror)
        WriteLog vbCrLf & "Comando SQL: " & sql
    End If
End Function

Private Function InsertStatoMonitor(conn As Object) As Boolean
    'record inserito per gestire la concorrenza con gli altri motori
    
    InsertStatoMonitor = False
    On Error GoTo err_handler
    
    Const MOTORE As String = "T"
    Dim sql As String
    Dim esito As String
    Dim descrizione As String
    Dim RecordsAffected As Long
    Dim strerror As String
    descrizione = "ELABORAZIONE..."
    esito = ""
   
    
    sql = "INSERT INTO StatoMonitor VALUES (" & _
          "'" & MOTORE & "', " & _
          "'" & CreaDataPerSQL(Now) & "', " & _
          "'" & esito & "', " & _
          "'" & Replace(Left(descrizione, 50), "'", "''") & "')"
    conn.execute sql
    
    
    InsertStatoMonitor = True
    Exit Function
    
err_handler:
    If Err.Number <> 0 Or conn.errors.Count > 0 Then
        strerror = "Error #" & Err.Number & vbCrLf & _
              "   (Description: " & Err.Description & ")" & vbCrLf & _
              "   (Source: " & Err.Source & ")" & vbCrLf
        strerror = strerror & "    (Error count : " & conn.errors.Count & ")" & vbCrLf
        strerror = strerror & "    (NativeError: " & conn.errors.Item(0).NativeError & ")" & vbCrLf
        strerror = strerror & "    (NE Description: " & conn.errors.Item(0).Description & ")" & vbCrLf
        strerror = strerror & "    (SQLState: " & conn.errors.Item(0).SQLState & ")" & vbCrLf
        WriteLog (strerror)
        WriteLog vbCrLf & "Comando SQL: " & sql
    End If
End Function
Private Function CreaDataPerSQL(ByVal data)
    If IsDate(data) Then
        CreaDataPerSQL = Year(data) & "/" & Right("00" & Month(data), 2) & "/" & Right("00" & Day(data), 2) & " " & Right("00" & Hour(data), 2) & ":" & Right("00" & Minute(data), 2) & ":" & Right("00" & Second(data), 2)
    Else
        CreaDataPerSQL = Year(Now) & "/" & Right("00" & Month(Now), 2) & "/" & Right("00" & Day(Now), 2) & " " & Right("00" & Hour(Now), 2) & ":" & Right("00" & Minute(Now), 2) & ":" & Right("00" & Second(Now), 2)
    End If
End Function


Private Sub Swap(ByRef s1 As String, ByRef s2 As String)
   Dim temp As String
   temp = s1
   s1 = s2
   s2 = temp
End Sub

Sub Main()
'-----------------------------------------------------------
'3-3-2005 :  MODIFICHE PER LETTURA TABELLA CONFERME BOOKING
' inserita su database ORACLE di Winscroll
'-----------------------------------------------------------
Dim flag_errore As String
flag_errore = "D" 'inizializzato a errore database

Const RICHIESTE_IN_LENGTH As Integer = 100
Const RICHIESTE_OU_LENGTH As Integer = 100

Const MAX_SAVE_TABLE_RETRY As Integer = 3
Const ERROR_KEY_DUPLICATED As Integer = 2627

Const NEWS_PERIOD As Integer = 7
Const INFINITE_PERIOD As String = "99991231"
Dim stepEseguito As String
Dim news As Object
Dim percorsoIni As String
Dim nomeIni As String
Dim transStarted
Dim COMTIServer As String
Dim sqlCM As String
Dim livello_errore As String
Dim codErr As String
Dim messag As String
Dim servizio As String
Dim data_iniz_rifer As Long
Dim ora_iniz_rifer As Integer
Dim data_fine_rifer As Long
Dim ora_fine_rifer As Integer
Dim anno_rich As Integer
Dim num_rich As Long
Dim vers_rich As Integer
Dim isn As Long
Dim data_inizio_trasm As Long
Dim ora_inizio_trasm As Long
Dim anno_start As Integer
Dim num_start As Long
Dim ver_start As Integer
Dim indic_segue_fine As String
Dim nr_richieste As Integer
Dim richieste() As RaiBook.M7375350_RICHIESTE
Dim ret As Long
Dim Leggi_CM_A7375351 As Boolean
Dim indic_acquisiz As String
Dim richieste_in As Object
Dim richieste_ou As Object
Dim rstCM1 As Object
Dim rstCM2 As Object
Dim sql As String
Dim Sql_update_sup As String

Dim data_inizio As String
Dim data_fine As String
Dim data_corrente As String
Dim date_format As String

Dim k As Integer
Dim i As Integer
Dim J As Integer
Dim retryCounter As Integer

Dim conn As Object
Dim rs As Object
Dim rs1 As Object
Dim rsp As Object
'Dim ErroreADO As ADODB.Error
Dim rstMC As Object
Dim rstctrl As Object

Dim prog As Integer
Dim DataOraInizio As String
Dim DataOraFine As String
Dim OraInizio As String
Dim OraFine As String
   
Dim rest As String
Dim ln As Integer
Dim pos As Integer
Dim EsitoPath As String
Dim NomeEsito As String
Dim msg As String
Dim indice As Integer
'Dim FileName As String
Dim flagEof_Rich As Boolean
Dim riaggancio_CM As Boolean

Dim trovatoRisorse
Dim strerror As String
Dim s As String
Dim params() As String
Dim esitoAbbin As Integer
Dim tipo_segnalazione_ponti As String
Dim tipo_segnalazione_richiesta As String
ErroreHost = 0
Set richieste_in = Nothing
Set richieste_ou = Nothing

On Error GoTo main_ErrorHandler
    
    If Left(GetDateFormat, 1) = "D" Then
        date_format = ITALIAN_DATE_FORMAT
    Else
        date_format = AMERICAN_DATE_FORMAT
    End If
    
   '''''' Parte nuova di decodifica parametri ''' rob 24/04/2003 '''''
    
    EnableLog = True ''' Rob 28/4/2003 : e' sempre abilitato il log
    percorsoIni = ""
    nomeIni = ""
    
    params = Split(Command$, ";")
    
    'Scorre il vettore dei parametri della linea di comando
    ' i parametri sono separati da ; (punto e virgola) e sono composti
    ' dal segno - (meno) seguito dall'id del parametro e subito il valore parametro
    '       es. -nprovaprog.ini     -n parametro nome file Ini
    '                               provaprog.ini e' il valore del parametro
    For i = 0 To UBound(params)
        s = Trim(params(i))
        If s = "" Then
            Exit For
        End If
        'esamino il parametro
        Select Case Left(s, 2)
            Case "-v"   'verbose = true
                EnableLog = True
            
            Case "-p"   'percorso file ini
                percorsoIni = Mid(s, 3)
                
            Case "-n"   ' nome file ini
                nomeIni = Mid(s, 3)
            
            Case "-i"   'data iniziale
                data_inizio = Mid(s, 3)
            
            Case "-f"   'data finale
                data_fine = Mid(s, 3)
            
        End Select
    Next
'Verifica la presenza delle date e imposta il formato YYYYMMDD
' se data_inizio non specificata la forza a data odierna
    'data odierna
    data_corrente = Format(Date, date_format)
    If IsDate(data_inizio) Then
       data_inizio = Format(data_inizio, ITALIAN_DATE_FORMAT)
    Else
    'imposta la data inizio a tre giorni prima della data odierna
        data_inizio = Format(DateAdd("d", -3, data_corrente), date_format)
       ''''''''''''''''''''''' forzato una prima volta  x caricare pregresso
       'data_inizio = "2004/01/01"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
    data_inizio = FormatDate(Format(data_inizio, ITALIAN_DATE_FORMAT))
' se data_fine non specificata la forza a data lontanissima 99991231
    If IsDate(data_fine) Then
       data_fine = FormatDate(Format(data_fine, ITALIAN_DATE_FORMAT))
    Else
       data_fine = INFINITE_PERIOD
    End If
   
'se percorsoIni non e' specificato imposta il percorso dell'applicativo
    If percorsoIni = "" Then
        percorsoIni = App.Path
    End If
'se percorsoIni non contiene il carattere \ lo concatena
    If Right(percorsoIni, 1) <> "\" Then
        percorsoIni = percorsoIni & "\"
    End If
'se nomeIni non e' specificato imposta il nome dell'applicativo
    If nomeIni = "" Then
        nomeIni = App.EXEName
    End If
'se nomeIni non contiene l'estensione .INI la concatena
    If UCase(Right(nomeIni, 4)) <> ".INI" Then
        nomeIni = nomeIni & ".INI"
    End If
'compone il nome completo del file INI
    AppIniFile = percorsoIni & nomeIni
    If Dir$(AppIniFile) = "" Then
'non trova il file INI
        Open App.Path & "\" & App.EXEName & ".007" For Output As #3 'Errore 1 manca il file ini
        Print #3, "Impossibile trovare o leggere il file " & AppIniFile
        Close #3
        End
    End If
   
   '**********************************************************************
   Dim RecordsAffected As Long
   Dim data_tx
   
    WriteLog "", True  'Serve per stampare la riga iniziale
   
    WriteLog "Inizio Elaborazione"
   'Inizialmente cancello tutti i file ESITO in app.path, che è il nome che si da di default
   'al file esito per gli errori che si verificano prima di leggere il file dei parametri:
   'quindi prima di leggere il nome da dare al file esito.
    Call CancellazioneFileEsito(App.Path & "\", "ESITO.0")
   'Lettura ESITO
    s = Space(255)
    ln = 255
    rest = GetPrivateProfileString("esito", "esito_file", "", s, ln, AppIniFile)
    FileEsito = Left(s, rest)
   '...  Cancello il file esito precedente  ...'
    If Right(FileEsito, 1) = "/" Or Right(FileEsito, 1) = "\" Then
        Open App.Path & "\" & App.EXEName & ".001" For Output As #3 'Errore 1 manca nome del file esito
        Close #3
        msg = "Manca il nome del file di esito del programma nel file ini"
        WriteLog (msg)
        End
    End If
    'Se ho la path del file esito,allora vado a cancellare il file esito precedente
    pos = InStr(1, FileEsito, "/")
    If pos = 0 Then
        pos = InStr(1, FileEsito, "\")
    End If
    If pos <> 0 Then
        EsitoPath = Percorso(FileEsito)
        NomeEsito = NomeEsitoDapas
    Else
        NomeEsito = FileEsito
        EsitoPath = App.Path & "\"
    End If
    FileEsito = EsitoPath & NomeEsito
    If EsitoPath = "$(AppPath)\" Then
        FileEsito = App.Path & "\" & NomeEsito
    Else
        FileEsito = EsitoPath & NomeEsito
    End If
'--------------- CODIFICA File di ESITO -----------------------------------
'FileEsito.000       Tutto OK, dati elaborati, inviare mail
'FileEsito.001       Impossibile aprire il file esito
'FileEsito.002       Impossibile aprire l'istanza per l'oggetto RaiBook
'FileEsito.003       Errore sulla connessione al db
'FileEsito.004       Warning dopo chiamata HOST
'FileEsito.005       Errore non previsto non dovuto al db SQL Server
'FileEsito.006       Errore dopo chiamata HOST
'FileEsito.007       File INI non trovato
'FileEsito.008       Un'istanza dell'applicazione e` gia` in esecuzione
'FileEsito.010       tutto OK, niente da fare, non inviare mail
'--------------------------------------------------------------------------
    Call CancellazioneFileEsito(EsitoPath, NomeEsito & ".0")
    
    'controllo che non ci sia una istanza della applicazione ancora attiva
    If App.PrevInstance = True Then
        Open FileEsito & ".008" For Output As #3 'Errore 8 istanza ancora attiva
        Close #3
        WriteLog "Una istanza precedente della applicazione e` ancora in esecuzione."
        WriteLog "Fine Programma KO"
        End
    End If
'**********************************************************************
    Dim strConn As String
    Dim ORA_strConn As String
    Dim vers_cm As Integer, vers_mr As Integer
    Dim esito As String
    esito = ""
On Error Resume Next
    strConn = GetDBConnection()
    If strConn = "" Then
        Open FileEsito & ".003" For Output As #3 'Errore stringa conn SQL non trovata
        Close #3
        WriteLog "Errore nella lettura stringa connessione DB SQL Server dal file INI"
        WriteLog "Fine Programma KO"
        End
    End If
    Set conn = CreateObject("ADODB.Connection")
    If Err.Number <> 0 Then
        Open FileEsito & ".003" For Output As #3
        Close #3
        WriteLog Err.Description
        WriteLog "Fine Programma KO"
        conn.Close
        End
    End If
    conn.Open strConn
    If Err.Number <> 0 Then
        Open FileEsito & ".003" For Output As #3
        Close #3
        WriteLog Err.Description
        WriteLog "Fine Programma KO"
        conn.Close
        End
    End If
    
Dim No_connessione_ora As Boolean
No_connessione_ora = False

    '3-3-2005 CONNESSIONE ORACLE  (ricerca la stringa di connessione Oracle nel file .ini e tenta la connessione)
    
    ORA_strConn = GetORADBConnection() 'preleva connessione oracle da .ini
    If ORA_strConn = "" Then
        Open FileEsito & ".003" For Output As #3 'Errore stringa conn ORACLE non trovata
        Close #3
        WriteLog "Errore nella lettura stringa connessione ORACLE dal file INI"
        WriteLog "Fine Programma KO"
        conn.Close
        End
    End If
    
    Set ORA_conn = CreateObject("ADODB.Connection")
    If Err.Number <> 0 Then
        Open FileEsito & ".003" For Output As #3
        Close #3
        WriteLog Err.Description
        WriteLog "Fine Programma KO"
        conn.Close
        End
    End If
      
    ORA_conn.Open ORA_strConn
    If Err.Number <> 0 Or ORA_conn.errors.Count > 0 Then
        No_connessione_ora = True
        'Open FileEsito & ".003" For Output As #3
        'Close #3
        WriteLog "La connessione non e' stata aperta: l'elaborazione prosegue con l'acquisizione dei collegamenti da host : " & Err.Description & ORA_conn.errors.Item(0).Description
        WriteLog "connessione : " & ORA_strConn
        Err.Clear
        'WriteLog "Fine Programma KO"
        'ORA_conn.Close
        'conn.Close
        'End
    Else
        'WriteLog "Eseguita connessione con la stringa seguente: " & ORA_strConn
    End If
    
    
'30-07-2004  gestione concorrenza con dll attivate dai bottoni su monitor
If Not CtrlStatoMonitor(conn) Then
        Open FileEsito & ".008" For Output As #3 'Errore 8 istanza ancora attiva
        Close #3
        WriteLog "Una istanza precedente della applicazione e` ancora in esecuzione."
        If Not SetStatoMonitor(conn, "M", "Applicazione in uso ") Then
            WriteLog "Non e' stato possibile effettuare aggiornamento su Tabella Monitor"
        End If
        WriteLog "Fine Programma KO"
        conn.Close
        If No_connessione_ora Then
        Else
         ORA_conn.Close
        End If
        End
End If

If Not InsertStatoMonitor(conn) Then 'gestione concorrenza fra motori
        Open FileEsito & ".003" For Output As #3
        Close #3
        If Not SetStatoMonitor(conn, "D", "Errore Database ") Then
            WriteLog "Non e' stato possibile effettuare aggiornamento su Tabella Monitor "
        End If
        WriteLog "Fine Programma KO"
        conn.Close
        If No_connessione_ora Then
        Else
         ORA_conn.Close
        End If
        End
End If

'Libero le eventuali richieste occupate.
    If Not DeleteOccupate(conn) Then
        Open FileEsito & ".003" For Output As #3
        Close #3
        WriteLog "Non e' stato possibile cancellare le richieste occupate  "
        If Not SetStatoMonitor(conn, "D", "Errore Database ") Then
            WriteLog "Non e' stato possibile effettuare aggiornamento su Tabella Monitor"
        End If
        If Not DeleteStatoMonitor(conn) Then 'gestione concorrenza fra motori
            WriteLog "Non e' stato possibile cancellare il record di elaborazione"
            WriteLog "Fine Programma KO"
            conn.Close
            If No_connessione_ora Then
            Else
             ORA_conn.Close
            End If
            End
        End If
        WriteLog "Fine Programma KO"
        conn.Close
        ORA_conn.Close
        End
    End If
    
    COMTIServer = GetServerCOMTI()
    If Trim(COMTIServer) = "" Then
        Set news = CreateObject("RaiBook.Novita.1")
    Else
        Set news = CreateObject("RaiBook.Novita.1", COMTIServer)
    End If
    If Err.Number <> 0 Then
        Open FileEsito & ".002" For Output As #3 'Errore 2 Impossibile aprire l'istanza per l'oggetto RaiBook
        Close #3
        msg = "Impossibile aprire l'istanza per l'oggetto RaiBook sul server " & COMTIServer & vbCrLf & Err.Description
        WriteLog (msg)
        If Not DeleteStatoMonitor(conn) Then 'gestione concorrenza fra motori
            WriteLog "Fine Programma KO"
            conn.Close
            ORA_conn.Close
            End
        End If
        WriteLog "Fine Programma KO"
        conn.Close
        ORA_conn.Close
        End
    End If
    RichAcq = 0
    servizio = "P" 'prima chiamata
    Dim richiesta_saltata As Boolean
    Dim descrizione_x_monitor As String
    richiesta_saltata = False
    descrizione_x_monitor = ""
    On Error GoTo main_ErrorHandler
Dim myMsg
    Do
        stepEseguito = "Inizio Loop"
        data_iniz_rifer = CLng(data_inizio)
        ora_iniz_rifer = 1
        data_fine_rifer = CLng(data_fine)
        ora_fine_rifer = 2400
        anno_rich = 0
        num_rich = 0
        vers_rich = 0
        ReDim richieste(14)
        For i = LBound(richieste) To UBound(richieste)
            With richieste(i)
                ReDim .A7375351_EL_TAB_RIS(14)
                ReDim .A7375351_EL_MR_MEZZO(3)
                .A7375351_CM_NOTE = String(50, " ")
                .A7375351_DATA_INOLTRO = 0
                .A7375351_ORA_INOLTRO = 0
            End With
        Next i
      
        ret = news.getNews(livello_errore, codErr, messag, servizio, data_iniz_rifer, ora_iniz_rifer, data_fine_rifer, ora_fine_rifer, anno_rich, num_rich, vers_rich, isn, data_inizio_trasm, ora_inizio_trasm, anno_start, num_start, ver_start, indic_segue_fine, nr_richieste, richieste)
        stepEseguito = "Chiamata al metodo getNews"
       
        servizio = "A"
        livello_errore = Trim$(livello_errore)
        
        If livello_errore = "E" And codErr = "E278" Then
        Else
            If livello_errore <> "" Then
                If Not SetStatoMonitor(conn, "E", "errore CICS") Then
                    WriteLog "Non e' stato possibile effettuare aggiornamento su Tabella Monitor "
                End If
            End If
        End If
        
        stepEseguito = "Aggiornamento dati x monitor"
        If livello_errore = "" Or livello_errore = "W" Then
            WriteLog "M7375350 : " & nr_richieste & " richieste lette"

            If livello_errore = "W" Then
                WriteLog "WARNING [" & codErr & "] chiamata modulo M7375350: " & messag
            End If
     
            For i = 0 To nr_richieste - 1
                Set richieste_in = news.NewRecordset("M7375400_RICHIESTE_IN")
                Set richieste_ou = news.NewRecordset("M7375400_RICHIESTE_OU")
                stepEseguito = "Chiamata al metodo news.NewRecordset(M7375400_RICHIESTE_OU)"
                RichAcq = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'controllo sugli orari : se errati la richiesta viene by-passata e viene fornita segnalazione sul log e sul monitor
        
        
        If richieste(i).A7375351_ORA_INIZIO > 2400 Then
            If Trim(richieste(i).A7375351_STATO_CIRCUITI) <> "SUP" Then
                WriteLog "La richiesta " & richieste(i).A7375351_ANNO_RICH & "/" & richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH & " ha il periodo errato"
                richiesta_saltata = True
                If descrizione_x_monitor = "" Then
                    descrizione_x_monitor = "(W) - Rich." & richieste(i).A7375351_ANNO_RICH & "/" & richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH & " bypassata x orario errato"
                End If
                GoTo prossima_richiesta
            Else
                GoTo prossima_richiesta
            End If
        End If
        
        If richieste(i).A7375351_ORA_FINE > 2400 Then
          If Trim(richieste(i).A7375351_STATO_CIRCUITI) <> "SUP" Then
            WriteLog "La richiesta " & richieste(i).A7375351_ANNO_RICH & "/" & richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH & " ha il periodo errato"
            richiesta_saltata = True
            If descrizione_x_monitor = "" Then
                descrizione_x_monitor = "(W) - Rich." & richieste(i).A7375351_ANNO_RICH & "/" & richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH & " bypassata x orario errato"
            End If
            GoTo prossima_richiesta
          Else
            GoTo prossima_richiesta
          End If
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'GIULIA 7-1-2003 verifica se la richiesta esiste gia' (in base al flag verra' deciso se fare insert o update, e viene tolto l'update dentro l'error handler)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' se e' vecchia ed e' versione > 1, e' da saltare
                data_tx = leggiDataHost(richieste(i).A7375351_DATA_INIZIO, richieste(i).A7375351_ORA_INIZIO)
                stepEseguito = "Lettura delle date dal tracciato host dell'oggetto COMTI"
                If DateDiff("d", data_corrente, data_tx) < 0 Then
                    stepEseguito = "Confronto tra data trasmissione e data corrente"
                    If richieste(i).A7375351_VER_RICH > 1 Then
                        stepEseguito = "Confronto su versione richiesta > 1"
                        WriteLog "Versione precedente della richiesta " & richieste(i).A7375351_ANNO_RICH & "/" & _
                            richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH & _
                            " (" & Format(data_tx, "dd/mm/yyyy") & ") gia` confermata"
                        GoTo prendi_in_carico
                    End If
                End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Set rstctrl = conn.execute("select * from Richiesta where Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " and Num_Rich = " & richieste(i).A7375351_NUM_RICH & " And Ver_Rich = " & richieste(i).A7375351_VER_RICH)
                stepEseguito = "Ricerca in archivio richieste su SQLServer"
                flagEof_Rich = rstctrl.EOF
                'se e` gia` presente e antecedente alla data del giorno salta alla prossima richiesta
                If (Not flagEof_Rich) Then
                    If DateDiff("d", data_corrente, rstctrl("data_ora_inizio")) < 0 Then
                        WriteLog "Richiesta " & richieste(i).A7375351_ANNO_RICH & "/" & _
                            richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH & _
                            " (" & Format(rstctrl("data_ora_inizio"), "dd/mm/yyyy") & ") gia` presente nel db SQLServer"
                        rstctrl.Close
                        Set rstctrl = Nothing
                        GoTo prendi_in_carico
                    End If
                    stepEseguito = "Confronto tra data inizio trasmissione (db) e data corrente"
                End If
                rstctrl.Close
                Set rstctrl = Nothing
                
                 'Controllo multiconcorrenza:
                ChkSalto = False
                'occupa la versione precedente (che verra' aggiornata con stato = s)
                If richieste(i).A7375351_VERSIONE_RIF <> 0 And richieste(i).A7375351_COD_TIPO_RICHIESTA = "V" And richieste(i).A7375351_VERSIONE_RIF < richieste(i).A7375351_VER_RICH Then
                    'WriteLog "select * from RichiesteOccupate where Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " and Num_Rich = " & richieste(i).A7375351_NUM_RICH & " And Ver_Rich = " & richieste(i).A7375351_VERSIONE_RIF
                    Set rstMC = conn.execute("select * from RichiesteOccupate where Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " and Num_Rich = " & richieste(i).A7375351_NUM_RICH & " And Ver_Rich = " & richieste(i).A7375351_VERSIONE_RIF)
                    EsitoMC = rstMC.EOF
                    rstMC.Close
                    Set rstMC = Nothing
                    If EsitoMC = False Then 'Occupata la deve saltare
                       ChkSalto = True
                       WriteLog "Richiesta " & richieste(i).A7375351_ANNO_RICH & "/" & richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VERSIONE_RIF & " occupata"
                    Else 'la deve bloccare
                      DataOraInizioMC = Year(Date) & "/" & Format(Month(Date), "00") & "/" & Format(Day(Date), "00") & " " & Format(Hour(Time), "00") & ":" & Format(Minute(Time), "00") & ":" & Format(Second(Time), "00")
                      conn.execute ("Insert into RichiesteOccupate (Anno_Rich,Num_Rich,Ver_Rich,Id,DataOraInizio) values (" & richieste(i).A7375351_ANNO_RICH & "," & richieste(i).A7375351_NUM_RICH & "," & richieste(i).A7375351_VERSIONE_RIF & ",99,'" & DataOraInizioMC & "')")
                      versione = richieste(i).A7375351_VERSIONE_RIF
                    End If
                End If
              ' ------------------
              ' occupa l'ultima versione se la richiesta esiste gia'
                If flagEof_Rich = False Then ' caso in cui la richiesta esiste gia' e viene quindi fatto il controllo multiutenza
                    Set rstMC = conn.execute("select * from RichiesteOccupate where Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " and Num_Rich = " & richieste(i).A7375351_NUM_RICH & " And Ver_Rich = " & richieste(i).A7375351_VER_RICH)
                    EsitoMC = rstMC.EOF
                    rstMC.Close
                    Set rstMC = Nothing
                    If EsitoMC = False Then 'Occupata la deve saltare
                       ChkSalto = True
                       WriteLog "Richiesta " & richieste(i).A7375351_ANNO_RICH & "/" & richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH & " occupata"
                    Else 'la deve bloccare
                       DataOraInizioMC = Year(Date) & "/" & Format(Month(Date), "00") & "/" & Format(Day(Date), "00") & " " & Format(Hour(Time), "00") & ":" & Format(Minute(Time), "00") & ":" & Format(Second(Time), "00")
                       conn.execute ("Insert into RichiesteOccupate (Anno_Rich,Num_Rich,Ver_Rich,Id,DataOraInizio) values (" & richieste(i).A7375351_ANNO_RICH & "," & richieste(i).A7375351_NUM_RICH & "," & richieste(i).A7375351_VER_RICH & ",99,'" & DataOraInizioMC & "')")
                       versione = richieste(i).A7375351_VER_RICH
                    End If
                End If
                  
                If ChkSalto Then
                'Se la richiesta e' occupata la deve saltare.
                    GoTo prossima_richiesta
                End If
' x debug                WriteLog "Tipo richiesta in partenza : " & richieste(i).A7375351_COD_TIPO_RICHIESTA

            tipo_segnalazione_richiesta = richieste(i).A7375351_COD_TIPO_RICHIESTA
            tipo_segnalazione_ponti = richieste(i).A7375351_COD_TIPO_SEGNALAZ_P
            'verifica se ci sono gia' ottimizzazioni:
            '           se SI lascia il valore
            '           se NO e il campo e' V lo riporta a N
                sql = "SELECT DISTINCT (num_ottim), anno_ottim,stato_risorsa,id_ver_rich " & _
                " FROM Risorse WHERE ID_Anno_Rich=" & richieste(i).A7375351_ANNO_RICH & " and " & _
                " ID_Num_Rich=" & richieste(i).A7375351_NUM_RICH & " and ID_Ver_Rich <= " & richieste(i).A7375351_VER_RICH & " and not(anno_ottim is null)"
                Set rstctrl = conn.execute(sql)
                If rstctrl.EOF Then
                    If richieste(i).A7375351_COD_TIPO_RICHIESTA = "V" Then
                        richieste(i).A7375351_COD_TIPO_RICHIESTA = "N"
                    End If
                End If
                rstctrl.Close
                Set rstctrl = Nothing
                
    ''''''''''''''''''''''''''''''''''''''''''''''
            Leggi_CM_A7375351 = False
            riaggancio_CM = False
          '----------------------------------------------------
          '  COLLEGAMENTO  CON WINSCROLL
          '----------------------------------------------------
          '3-3-2005
          'LETTURA VISTA ORACLE RELATIVA AI COLLEGAMENTI MOBILI
          If Trim(richieste(i).A7375351_STATO_PONTI) = "CON" Then
            On Error Resume Next
            If No_connessione_ora Then
                Leggi_CM_A7375351 = True
                GoTo Prosegui_elaborazione
            End If
'se ci sono dei problemi con la nuova tabella scommentare le due istruzioni qui sotto
'Leggi_CM_A7375351 = True
'GoTo Prosegui_elaborazione
            'in caso di errori o per non trovato sulla tabella utilizza i dati
            'dell'area A7375351 (che continua ad essere impostata)
            'e si evita di fermare l'elaborazione
            'record C1
            sqlCM = "SELECT * " & _
            " FROM CONFERMEBOOKING WHERE Anno_richiesta =" & richieste(i).A7375351_ANNO_RICH & " and " & _
            " Numero_richiesta = " & richieste(i).A7375351_NUM_RICH & " and Versione_richiesta = " & richieste(i).A7375351_VER_RICH & " and progressivo = 1"
            
            Set rstCM1 = CreateObject("ADODB.Recordset")
            'connessione Oracle
            rstCM1.ActiveConnection = ORA_conn
            rstCM1.CursorType = 3
            rstCM1.Open sqlCM
             
            If Err <> 0 Then
                Open FileEsito & ".003" For Output As #3
                Close #3
                WriteLog Err.Description & " sql : " & sqlCM
                Leggi_CM_A7375351 = True
                GoTo Prosegui_elaborazione
            End If
        
          
            If rstCM1.EOF Then
                'If richieste(i).A7375351_VER_RICH > 1 Then
                '   If tipo_segnalazione_ponti = "N" And tipo_segnalazione_richiesta = "V" Then 'la versione e' cambiata per varianti non relative ai ponti
                '        If Riaggancia_Versione_prec(conn, richieste(i).A7375351_ANNO_RICH, richieste(i).A7375351_NUM_RICH, richieste(i).A7375351_VER_RICH, richieste(i).A7375351_VERSIONE_RIF) Then
                '            'riscrive il record  leggendo dalla versione precedente
                '            'WriteLog "le condizioni per il riaggancio sono presenti "
                '            riaggancio_CM = True
                '        Else
                '            'errore nel riaggancio o non trovata versione precedente
                '            Leggi_CM_A7375351 = True
                '            WriteLog "Record di conferma CM non trovato - sql : " & sqlCM
                '            GoTo Prosegui_elaborazione
                '        End If
                '   Else
                '        Leggi_CM_A7375351 = True
                '        WriteLog "Record di conferma CM non trovato - sql : " & sqlCM
                '        GoTo Prosegui_elaborazione
                '   End If
                'Else
                    Leggi_CM_A7375351 = True
                    WriteLog "Record di conferma CM non trovato: - " & richieste(i).A7375351_ANNO_RICH & " - " & richieste(i).A7375351_NUM_RICH
                    GoTo Prosegui_elaborazione
                'End If
            End If
            Dim no_C2 As Boolean
            no_C2 = False
            
            'record C2    (PUO' ESSERE PRESENTE O MENO)
            sqlCM = "SELECT * " & _
            " FROM CONFERMEBOOKING WHERE Anno_richiesta =" & richieste(i).A7375351_ANNO_RICH & " and " & _
            " Numero_richiesta = " & richieste(i).A7375351_NUM_RICH & " and Versione_richiesta = " & richieste(i).A7375351_VER_RICH & " and progressivo = 2"
         
            Set rstCM2 = CreateObject("ADODB.Recordset")
            ' connessione ORACLE
            rstCM2.ActiveConnection = ORA_conn
            rstCM2.CursorType = 3
            rstCM2.Open sqlCM
            
            If Err <> 0 Then
                Open FileEsito & ".003" For Output As #3
                Close #3
                WriteLog Err.Description & " sql : " & sqlCM
                GoTo Prosegui_elaborazione
                Leggi_CM_A7375351 = True
                GoTo Prosegui_elaborazione
            End If
            If rstCM2.EOF Then
                no_C2 = True
            End If
          End If 'ponti confermati

Prosegui_elaborazione:
On Error GoTo main_ErrorHandler

    'INSERIMENTO O UPDATE DELLA RICHIESTA
                If flagEof_Rich = True Then
                'Nuova richiesta (o nuova versione)
                    sql = GetInsertStatementForRichiesta(richieste(i))
                    vers_cm = 1
                    vers_mr = 1
                Else
                'Variante di richiesta esistente
                    sql = GetUpdateStatementForRichiesta(richieste(i))
                End If
        '////   Inizio TRANSAZIONE ///////  dal 20-03-2003 viene eseguita per ogni richiesta
                transStarted = conn.BeginTrans()
                
                If richieste(i).A7375351_COD_TIPO_RICHIESTA = "V" Then
                'verifica se la risorsa ha ottimizzazioni e fonogrammi ed eventualmente
                'imposta il segnalino di allarme su ottim_fono
                    If Not risorsaConOttimizzazioni(conn, richieste(i).A7375351_ANNO_RICH, richieste(i).A7375351_NUM_RICH, richieste(i).A7375351_VER_RICH) Then
                        conn.RollbackTrans
                        transStarted = 0
                        WriteLog "Errore su richiesta: " & richieste(i).A7375351_ANNO_RICH & "/" & richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH
                        WriteLog "RollbackTrans SQLServer"
                        'Libero le eventuali richieste occupate.
                        conn.execute ("delete from RichiesteOccupate where Id = 99")
                         'elimina record elaborazione
                        conn.execute ("Delete from StatoMonitor where ID_MOTORE = 'T'")
                        conn.Close
                        Set news = Nothing
                        Open FileEsito & ".003" For Output As #3 'Errore 3 errore sulla connessione al db
                        Close #3
                        WriteLog "Fine Programma KO"
                       End
                    End If
                End If
                'INSERISCO o AGGIORNO LA RIchiesta sul Database SQL
                conn.execute sql
                
                RichAcq = RichAcq + 1
                retryCounter = 1

'----------------------  NUOVA GESTIONE CAMPI COLLEGAMENTI MOBILI
'----------------------  PRESI DA TABELLA WS O DA SOLITA AREA DI COMUNICAZIONE

        '------------------------------------------------------------------------------------------------------------------------------------
          If Trim(richieste(i).A7375351_STATO_PONTI) = "CON" Then ' ------------
            If riaggancio_CM = False Then
             If Not Leggi_CM_A7375351 Then 'trovato record di conferma SU DB Winscroll
               WriteLog "elaborazione dati tabella Conferme Booking : - " & richieste(i).A7375351_ANNO_RICH & " - " & richieste(i).A7375351_NUM_RICH
               'confronta i dati con una eventuale
               'versione gia' memorizzata nella tabella COLLEGAMENTI MOBILI
               'e scrive eventualmente la versione successiva
                  'ricava la versione piu' recente
                   vers_cm = getVersione("CM", conn, richieste(i).A7375351_ANNO_RICH, richieste(i).A7375351_NUM_RICH, richieste(i).A7375351_VER_RICH)
                   WriteLog "RICERCA ultima versione"
                  'confronta con la nuova tabella Winscroll
                      stepEseguito = "prima dell'innesco della function di controllo dei collegamenti mobili"
                   If Not isStessiDati_NewCollegMobili(conn, richieste(i).A7375351_ANNO_RICH, richieste(i).A7375351_NUM_RICH, richieste(i).A7375351_VER_RICH, vers_cm, rstCM1, no_C2, rstCM2) Then
                      WriteLog "Confronto dati "
                      vers_cm = vers_cm + 1
                      stepEseguito = "prima dell'inserimento di collegamenti mobili"
                      sql = getInsertNewCollegamentoMobile(richieste(i).A7375351_ANNO_RICH, richieste(i).A7375351_NUM_RICH, richieste(i).A7375351_VER_RICH, vers_cm, rstCM1, rstCM2)
                      If sql = "" Then
                        Leggi_CM_A7375351 = True
                      Else
                        On Error Resume Next
                        conn.execute sql
                      End If
                      If Err.Number <> 0 Or conn.errors.Count > 0 Then
                        Leggi_CM_A7375351 = True
                        WriteLog "Si e' verificato un errore nell'inserimento dei nuovi dati dei CM : " & Err.Number & " - " & Err.Description & " - " & sql
                      Else
                        WriteLog "Inserimento effettuato"
                      End If
                   End If
                   rstCM1.Close
                   rstCM2.Close
                   Set rstCM1 = Nothing
                   Set rstCM2 = Nothing

            End If  '
           End If
         End If  '
On Error GoTo main_ErrorHandler
  '------------------------------------------------------------------------------------------------------------------------------------

If Leggi_CM_A7375351 Then 'vecchia gestione (Record non trovato su tabella o errore)
  If riaggancio_CM = False Then
'''''''''''''''
        'inserisco eventuali coll. mobili relativi alla richiesta
                If Trim(richieste(i).A7375351_STATO_PONTI) = "CON" Then
                    vers_cm = getVersione("CM", conn, richieste(i).A7375351_ANNO_RICH, richieste(i).A7375351_NUM_RICH, richieste(i).A7375351_VER_RICH)
                    If Not isStessiDati_CollegMobili(richieste(i), conn, vers_cm) Then
                        vers_cm = vers_cm + 1
                        sql = getInsertCollegamentoMobile(vers_cm, richieste(i))
                        If sql <> "" Then
                            'WriteLog "stringa : " & sql
                            conn.execute sql 'conn.error non e' intercettato dalla handle
                            If Err.Number <> 0 Or conn.errors.Count > 0 Then
                                GoTo main_ErrorHandler
                            End If
                        End If
                    End If
                End If
    End If
End If
'GIULIA MODIFICA DEL 5-12-2002
'Update delle richieste che sono state superate da una nuova versione
' modific. il 15/9/2003
' se esiste una versione precedente va aggiornata a SUPERATA
                If (richieste(i).A7375351_VERSIONE_RIF < richieste(i).A7375351_VER_RICH) And richieste(i).A7375351_VERSIONE_RIF > 0 Then
                    Sql_update_sup = GetUpdateSuperateForRichiesta(richieste(i))
                    conn.execute Sql_update_sup
                    Sql_update_sup = GetUpdateForRisorseSuperate(richieste(i))
                    conn.execute Sql_update_sup, RecordsAffected
                End If
            'Fine modifica Giulia
                
                If Trim(richieste(i).A7375351_STATO_MEZZI) = "CON" Then
                    vers_mr = getVersione("MR", conn, richieste(i).A7375351_ANNO_RICH, richieste(i).A7375351_NUM_RICH, richieste(i).A7375351_VER_RICH)
                    If Not isStessiDati_Mezzi(richieste(i), conn, vers_mr) Then
                        vers_mr = vers_mr + 1
                          ' ciclo di inserimento dei mezzi
                        For J = LBound(richieste(i).A7375351_EL_MR_MEZZO) To UBound(richieste(i).A7375351_EL_MR_MEZZO)
                            retryCounter = 1
                            If richieste(i).A7375351_EL_MR_MEZZO(J).A7375351_MR_MEZZO = "" Then
                                Exit For
                            End If
                            sql = GetInsertStatementForMezzi_Ripresa(richieste(i), J, vers_mr)
                            conn.execute sql
                        Next J
                    End If
                End If
' x debug                       WriteLog "Inizio insert risorse:"
                      ' cancello le eventuali risorse che non sono state ottimizzate
                sql = "DELETE FROM Risorse " & _
                      "WHERE ID_Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " AND " & _
                      "ID_Num_Rich = " & richieste(i).A7375351_NUM_RICH & " AND " & _
                      "ID_Ver_Rich = " & richieste(i).A7375351_VER_RICH & " AND " & _
                      "Anno_Ottim is Null AND Num_Ottim is Null"
                conn.execute sql
' x debug                       WriteLog "Esito cancella risorse: OK = " & Sql
                      'cerca il progressivo maggiore
                prog = 0
                sql = "SELECT MAX(Progressivo) AS prog FROM Risorse " & _
                      "WHERE ID_Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " AND " & _
                      "ID_Num_Rich = " & richieste(i).A7375351_NUM_RICH & " AND " & _
                      "ID_Ver_Rich = " & richieste(i).A7375351_VER_RICH
                Set rs = conn.execute(sql)
                If Not IsNull(rs("prog")) Then
                    prog = rs("prog")
                End If
                rs.Close
                '
                '---------------------------------------------------
                ' GIULIA GESTIONE ANNULLAMENTI
                '03-07-2003  Per poter distinguere gli annullamenti
                'precedenti bisogna:
                '-------------------------- 1 ----------------------
                'impostare indic-annullata con un altro carattere  es. x
                'escludendo dall'update quelle che hanno gia'
                'la colonna indic_annulla =  "A"
                '----------------------------2-----------------------
                'nel caso di match riuscito
                'bisogna impostare indic-annullata a null
                'escludendo dall'update quelle che hanno gia'
                'la colonna indic_annulla =  "A"
                '----------------------------3-----------------------
                'alla fine bisogna impostare indic-annullata ad A
                'in tutte le risorse che hanno la colonna
                'indic_annulla =  "X"
                '---------------------------------------------------
                '
                sql = "UPDATE Risorse SET Indic_Annullata = 'X' " & _
                      "WHERE ID_Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " AND " & _
                      "ID_Num_Rich = " & richieste(i).A7375351_NUM_RICH & " AND " & _
                      "ID_Ver_Rich = " & richieste(i).A7375351_VER_RICH & " AND " & _
                      "(Indic_Annullata is null OR Indic_Annullata <> 'A')"
                conn.execute sql
                      
                For J = LBound(richieste(i).A7375351_EL_TAB_RIS) To UBound(richieste(i).A7375351_EL_TAB_RIS)
                    retryCounter = 1
                    If richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_COD = "" Then
                        Exit For
                    End If
                    DataOraInizio = Right(String(4, "0") & richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_ORA_INIZ, 4)
                    OraInizio = convertiOre24(DataOraInizio)
                    DataOraInizio = richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_DATA_INIZ & " " & OraInizio
                    DataOraFine = Right(String(4, "0") & richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_ORA_FINE, 4)
                    OraFine = convertiOre24(DataOraFine)
                    DataOraFine = richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_DATA_FINE & " " & OraFine
                    sql = "SELECT * FROM Risorse " & _
                          "WHERE ID_Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " AND " & _
                          "ID_Num_Rich = " & richieste(i).A7375351_NUM_RICH & " AND " & _
                          "ID_Ver_Rich = " & richieste(i).A7375351_VER_RICH & " AND " & _
                          "Cod_Risorsa = '" & richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_COD & "' AND " & _
                          "Quantita = '" & richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_QTA & "' AND " & _
                          "Data_Ora_Inizio = '" & DataOraInizio & "' AND " & _
                          "Data_Ora_Fine = '" & DataOraFine & "' AND Indic_Annullata <> 'A'"
' x debug                            WriteLog " select risorsa: " & Sql
                    Set rs1 = conn.execute(sql)
                    trovatoRisorse = Not rs1.EOF
                    rs1.Close
                    If Not trovatoRisorse Then
                        sql = GetInsertStatementForRisorse(richieste(i), prog, J)
                        conn.execute (sql)
                        prog = prog + 1
                    Else 'c'e' la risorsa
                         '-----------------------giulia---------------------------
                         ' aggiorno la colonna indic_annulla con "NULL"
                        sql = "UPDATE Risorse SET Indic_Annullata = NULL " & _
                              "WHERE ID_Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " AND " & _
                              "ID_Num_Rich = " & richieste(i).A7375351_NUM_RICH & " AND " & _
                              "ID_Ver_Rich = " & richieste(i).A7375351_VER_RICH & " AND " & _
                              "Cod_Risorsa = '" & richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_COD & "' AND " & _
                              "Quantita = '" & richieste(i).A7375351_EL_TAB_RIS(J).A7375351_EL_QTA & "' AND " & _
                              "Data_Ora_Inizio = '" & DataOraInizio & "' AND " & _
                              "Data_Ora_Fine = '" & DataOraFine & "' AND " & _
                              "Indic_annullata <> 'A' "
                        conn.execute sql
                    End If
                         '-----------------------giulia fine -----------------------
                Next J
                ' tentativo di inserire piccolo ritardo per capire se l'errore
                ' di deadlock che si verifica una tantum e` dovuto alla sequenza di update
                ' su risorse se non si entra nel ciclo  for J = LBound(r ...
                Dim aa As Integer, bb As Double
                For aa = 1 To 100
                    bb = aa * 3.5 / 7
                Next
                
                ' tutte quelle rimaste sono da annullare
                sql = "UPDATE Risorse SET Indic_Annullata = 'A' " & _
                      "WHERE ID_Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " AND " & _
                      "ID_Num_Rich = " & richieste(i).A7375351_NUM_RICH & " AND " & _
                      "ID_Ver_Rich = " & richieste(i).A7375351_VER_RICH & " AND " & _
                      "Indic_Annullata = 'X'"
                conn.execute sql

prendi_in_carico:
                indic_acquisiz = "S"
                '******************* 17/9/2003 modifica sospesa, da rivedere
                'If richieste(i).A7375351_COD_TIPO_RICHIESTA = "V" Then
                '    esitoAbbin = AbbinaOttimPrec(conn, richieste(i).A7375351_ANNO_RICH, richieste(i).A7375351_NUM_RICH, richieste(i).A7375351_VER_RICH, msg)
                '    Select Case esitoAbbin
                '        Case 1  'ha abbinato
                '            indic_acquisiz = "U"
                '            WriteLog "Ri-abbinamento: richiesta ri-abbinata"
                        
                '        Case 2  'non ha abbinato perche` ci sono differenze
                '            WriteLog "Ri-abbinamento: la richiesta non e` stata ri-abbinata perche` esistono differenze tra i dati delle versioni"
                        
                '        Case -1 'si e` verificato un errore
                '            WriteLog "Errore Ri-abbinamento: " & msg
                '            conn.RollbackTrans
                '            transStarted = 0
                '            WriteLog "RollbackTrans SQLServer"
                '            GoTo prossima_richiesta
                            
                '    End Select
                'End If
                '+++++++++++++++++++
                     'inserimento della richiesta nel recordset per la conferma dell'acquisizione
                richieste_in.AddNew
                richieste_in!M7375400_ANNO_IN = richieste(i).A7375351_ANNO_RICH
                richieste_in!M7375400_NUM_IN = richieste(i).A7375351_NUM_RICH
                richieste_in!M7375400_VERS_IN = richieste(i).A7375351_VER_RICH
                richieste_in.Update
                'Libero la richiesta occupata.
                conn.execute ("delete from RichiesteOccupate where Anno_Rich = " & richieste(i).A7375351_ANNO_RICH & " and Num_Rich = " & richieste(i).A7375351_NUM_RICH & " and Ver_Rich = " & versione & " and Id = 99")
      
                'INIZIO GIULIA --20/03/2003-- CHIAMATA DEL M7375400 ad ogni richiesta
                For k = 1 To 99   'fisso 99 perche' ne riceve una per volta
                    richieste_in.AddNew
                    richieste_in!M7375400_ANNO_IN = 0
                    richieste_in!M7375400_NUM_IN = 0
                    richieste_in!M7375400_VERS_IN = 0
                    richieste_in.Update
                Next k
                richieste_in.movefirst
                'inizializzo il recordset di output
                For k = 1 To RICHIESTE_OU_LENGTH
                    richieste_ou.AddNew
                    richieste_ou!M7375400_ANNO_OU = 0
                    richieste_ou!M7375400_NUM_OU = 0
                    richieste_ou!M7375400_VERS_OU = 0
                    richieste_ou!M7375400_LIVERR_RICH = ""
                    richieste_ou!M7375400_CODERR_RICH = ""
                    richieste_ou.Update
                Next k
                richieste_ou.movefirst
                '--------------------------------
                'CHIAMATA
                '--------------------------------
                ret = news.NewsAck(indic_acquisiz, richieste_in, livello_errore, codErr, messag, richieste_ou)
                livello_errore = Trim$(livello_errore)
                ' GESTIONE ERRORI
                If livello_errore = "" Or livello_errore = "W" Then
                    WriteLog "M7375400 OK: " & richieste(i).A7375351_ANNO_RICH & "/" & richieste(i).A7375351_NUM_RICH & "/" & richieste(i).A7375351_VER_RICH
                    If transStarted <> 0 Then
                        conn.CommitTrans 'DEVE CONSOLIDARE ANCHE PER ERRORE = WARNING
                                    'PERCHE' ANCHE IL MODULO IN QUESTO CASO HA CONSOLIDATO
                        transStarted = 0
                        WriteLog "CommitTrans SQLServer"
                    End If
                   If livello_errore = "W" Then
                       WriteLog "WARNING [" & codErr & "] chiamata modulo M7375400: " & messag
                   End If
                Else
                    If transStarted <> 0 Then
                        conn.RollbackTrans 'LIVELLO ERRORE = A  o  E
                        transStarted = 0
                        WriteLog "RollbackTrans SQLServer"
                    End If
                    WriteLog "ERRORE [" & codErr & "] chiamata modulo M7375400: " & messag
                    ErroreHost = ErroreHost + 1
                    If codErr <> "E160" Then ' filtrato errore di richiesta occupata
                       If ErroreHost >= 1 Then
                          Open FileEsito & ".006" For Output As #3 'Errore 6 errore HOST
                          Close #3
                       End If
                    Else
                       Open FileEsito & ".000" For Output As #3
                       Close #3
                    End If
                    richieste_ou.movefirst
                    Do Until richieste_ou.EOF
                        If richieste_ou!M7375400_LIVERR_RICH <> "" Then
                        ' 2002/10/02 gestione degli errori e dei warning
                            WriteLog vbTab & "richiesta (anno - num - ver) " & richieste_ou!M7375400_ANNO_OU & " - " & richieste_ou!M7375400_NUM_OU & " - " & richieste_ou!M7375400_VERS_OU & " non acquisita da sistema per errore [" & codErr & "]: " & messag
                        End If
                        richieste_ou.movenext
                    Loop
                End If
          '---------------------------------------------------------------
          'FINE GIULIA 20-03-2003
          '---------------------------------------------------------------
prossima_richiesta:
                Set richieste_in = Nothing
                Set richieste_ou = Nothing
            Next i
         
      ElseIf livello_errore = "E" Then 'ELSE DEL LIVELLO ERRORE DEL MODULO M7375350
         If codErr = "E278" Then
            'non ci sono novità (questo caso non viene considerato errore)
            Exit Do
         Else
            '///////////
            WriteLog "ERRORE [" & codErr & "] chiamata modulo M7375350: " & messag
            WriteLog "Fine chiamata modulo M7375350"
            ErroreHost = ErroreHost + 1
            If ErroreHost >= 1 Then
               Open FileEsito & ".006" For Output As #3 'Errore 6 errore HOST
               Close #3
            End If
               flag_errore = "C"
               GoTo termina_conErrore
         End If
      Else
         '///////////
         WriteLog "ERRORE [" & codErr & "] chiamata modulo M7375350: " & messag
         WriteLog "Fine chiamata modulo M7375350"
           ErroreHost = ErroreHost + 1
           If ErroreHost >= 1 Then
              Open FileEsito & ".006" For Output As #3 'Errore 6 errore HOST
              Close #3
           End If
             flag_errore = "C"
             GoTo termina_conErrore
      End If
   
   Loop Until indic_segue_fine = "F"
   
   'aggiorna Monitor con esito positivo
   
    If Not richiesta_saltata Then
        If Not SetStatoMonitor(conn, "", "") Then
            WriteLog "Non e' stato possibile effettuare aggiornamento su Tabella Monitor "
        End If
    End If
   
   
   'Libero le eventuali richieste occupate.
   conn.execute ("delete from RichiesteOccupate where Id = 99")
    'elimina record elaborazione
   conn.execute ("Delete from StatoMonitor where ID_MOTORE = 'T'")
   If No_connessione_ora Then
   Else
    ORA_conn.Close
   End If
   Set news = Nothing
   If richiesta_saltata Then
            If Not SetStatoMonitor(conn, "W", descrizione_x_monitor) Then
                 WriteLog "Non e' stato possibile effettuare aggiornamento su Tabella Monitor "
            End If
            Open FileEsito & ".020" For Output As #3 'Nessun Errore ma e' stata by-passata una richiesta
            Close #3
   Else
    If ErroreHost = 0 Then
        If nr_richieste > 0 Then
            Open FileEsito & ".000" For Output As #3 'Nessun Errore e ricevuto richieste
            Close #3
        Else
            Open FileEsito & ".010" For Output As #3 'Nessun Errore e nessuna richiesta
            Close #3
        End If
    End If
    End If
    conn.Close
    WriteLog "Fine Programma OK"

   Exit Sub

main_ErrorHandler:
    strerror = "Error #" & Err.Number & vbCrLf & _
               "   (Description: " & Err.Description & ")" & vbCrLf & _
               "   (Source: " & Err.Source & ")" & vbCrLf & _
               "   (Ultima operazione eseguita: " & stepEseguito & ")" & vbCrLf
    If conn.errors.Count > 0 Then
         strerror = strerror & "    (Error count : " & conn.errors.Count & ")" & vbCrLf
         strerror = strerror & "    (NativeError: " & conn.errors.Item(0).NativeError & ")" & vbCrLf
         strerror = strerror & "    (NE Description: " & conn.errors.Item(0).Description & ")" & vbCrLf
         strerror = strerror & "    (SQLState: " & conn.errors.Item(0).SQLState & ")" & vbCrLf
    End If
    WriteLog (strerror)
    If conn.errors.Count > 0 Then
        WriteLog vbCrLf & "Comando SQL: " & sql
        '*************************************************************
        Open FileEsito & ".003" For Output As #3 'Errore 3 errore sulla connessione al db
        Close #3
        '*************************************************************
    Else
      '*************************************************************
       Open FileEsito & ".005" For Output As #3 'Errore 5 errore non aspettato non dovuto al db
       Close #3
    End If
termina_conErrore:
    'Dim Descrizione_errore As String
    On Error Resume Next
    If transStarted <> 0 Then
        'Rollback
            If Not funct_rollback(conn) Then 'gestione concorrenza fra motori
                WriteLog "Fine Programma KO"
                End
            End If
        transStarted = 0
    End If
'Elimina record elaborazione //
    If Not DeleteStatoMonitor(conn) Then 'gestione concorrenza fra motori
        WriteLog "Fine Programma KO"
        End
    End If
'Libero le eventuali richieste occupate //
    If Not DeleteOccupate(conn) Then 'gestione concorrenza fra motori
        WriteLog "Fine Programma KO"
        End
    End If
'Scrittura monitor // per scrivere che l'elaborazione e' andata in errore
    If flag_errore = "C" Then
        If Not SetStatoMonitor(conn, "E", "Errore CICS ") Then
            WriteLog "E' stato impossibile l'aggiornamento sulla tabella Monitor,  l'elaborazione prosegue"
        End If
    Else
        If Not SetStatoMonitor(conn, "D", "Errore Database ") Then
            WriteLog "E' stato impossibile l'aggiornamento sulla tabella Monitor,  l'elaborazione prosegue"
        End If
    End If
    conn.Close
    If No_connessione_ora Then
    Else
     ORA_conn.Close
    End If
    Set news = Nothing
    WriteLog "Fine Programma KO"
End Sub

Private Function GetLogInfo(ByVal param As String) As String
    Dim s As String
    Dim N As Integer
    Dim i As Integer
    Dim ret As Long
   
    s = Space(255)
    N = Len(s)
    ret = GetPrivateProfileString("log", param, "", s, N, AppIniFile)
    s = Mid(s, 1, ret)
      
    GetLogInfo = s
End Function

Function sostituisciAppPath(ByVal mypath As String) As String
Dim pos
    pos = InStr(1, mypath, "$(AppPath)")
    
    If pos > 0 Then
        sostituisciAppPath = App.Path & Mid(mypath, pos + Len("$(AppPath)"))
    Else
        sostituisciAppPath = mypath
    End If

End Function

Private Function WriteLog(ByVal msg As String, Optional InsertLine As Boolean = False, Optional FileName As String = "") As Boolean
    Dim i As Integer
    Dim pathstorico
    Dim logFileName As String
    
    If EnableLog Then
        On Error GoTo WriteLog_ErrorHandler
        'If FileName = "" Then FileName = GetLogFileName
        If FileName = "" Then
            logFileName = sostituisciAppPath(GetLogInfo("log_file"))
        Else
            logFileName = FileName
        End If
   'gestione del log giorno precedente
        If Dir(logFileName) <> "" Then
            'esiste gia' un file di log, verifico le date
            Dim dt1, dt2
            
            pathstorico = sostituisciAppPath(GetLogInfo("path_storico_log"))
            'If pathstorico = "" Then pathstorico = App.Path
            If Right(pathstorico, 1) <> "\" Then pathstorico = pathstorico & "\"
            
            dt1 = FormatDate(Format(FileDateTime(logFileName), ITALIAN_DATE_FORMAT))    'data del file in fmt YYYYMMDD
            dt2 = FormatDate(Format(Now, ITALIAN_DATE_FORMAT))                       'data corrente in fmt YYYYMMDD
            If dt1 <> dt2 Then
                'copia il file del giorno precedente
'                FileCopy FileName, App.Path & "\" & App.EXEName & dt1 & ".log"
                FileCopy logFileName, pathstorico & App.EXEName & dt1 & ".log"
                ' lo cancella
                Kill logFileName
            End If
        End If
        
        i = FreeFile
        Open logFileName For Append Access Write As #i
        
        If InsertLine Then
           Print #i, String(80, "-")
        Else
           Print #i, Format(Now, "DD/MM/YYYY hh:mm:ss") & " - " & msg
        End If
                 
        Close #i
    End If
    WriteLog = True
    Exit Function
   
WriteLog_ErrorHandler:
    WriteLog = False
End Function

Function Percorso(EsitoPath) As String
    Dim NomeEsitoR, carattere As String
    Dim i As Integer
    Percorso = ""
    For i = Len(EsitoPath) To 0 Step -1
        carattere = Mid(EsitoPath, i, 1)
        If carattere = "\" Or carattere = "/" Then
           Percorso = Left(EsitoPath, i)
           NomeEsitoDapas = Right(EsitoPath, Len(EsitoPath) - Len(Percorso))
           Exit Function
        Else
           NomeEsitoR = carattere & NomeEsitoR
        End If
    Next
End Function

Sub CancellazioneFileEsito(Percorso, Nome)
    Dim mypath, MyName As String
    mypath = Percorso    ' Imposta il percorso.
    MyName = Dir(mypath, vbDirectory)   ' Recupera la prima voce.
    Do While MyName <> ""   ' Avvia il ciclo.
       If (GetAttr(mypath & MyName) And vbDirectory) <> vbDirectory Then 'Se diverso da una directory
          If Left(UCase(MyName), Len(Nome)) = UCase(Nome) Then
             Kill Percorso & MyName
          End If
       End If
       MyName = Dir   ' Legge la voce successiva.
    Loop
End Sub

Private Function getVersione(ByVal tabella As String, cn As Object, ByVal anno As Integer, ByVal num As Long, ByVal ver As Integer) As Long
Dim sql, ret, rs

getVersione = 0
On Error GoTo err_getVersione
    Select Case tabella
        Case "CM"
            sql = "Select Max(Versione_CM) as vers From CollegamentiMobili Where " & _
            "anno_rich_cm=" & anno & " and num_rich_cm=" & num & " and ver_rich_cm=" & ver
        
        Case "MR"
            sql = "Select Max(Versione_MR) as vers From MezziRipresa Where " & _
            "id_anno_rich=" & anno & " and id_num_rich=" & num & " and id_ver_rich=" & ver
    End Select
    
    Set rs = cn.execute(sql)
       
    getVersione = IIf(IsNull(rs(0)), 0, rs(0))
    
    rs.Close
    
    Exit Function

err_getVersione:
    WriteLog "Errore nella funzione getVersione: " & Err.Description & vbCrLf & "SQL : " & sql

End Function
Private Function Riaggancia_Versione_prec(cn As Object, ByVal anno As Integer, ByVal num As Long, ByVal ver As Integer, ByVal ver_prec As Integer) As Boolean
Riaggancia_Versione_prec = False
On Error GoTo err_riaggancio

Dim ult_ver_cm As Integer
Dim rs
Dim CM_Ora_Iniz_PM As String
Dim CM_Ora_Fine_PM As String
Dim C2_Ora_Iniz_PM As String
Dim C2_Ora_Fine_PM As String
Dim sql As String, txp1, txp2, loc1, loc2
Dim sist1, sist2, note1, note2, audio1_1, audio1_2, audio2_1, audio2_2, tel1, tel2
Dim satellite1 As String
Dim Cod_satellite1 As String
Dim NRO_Cod_satellite1 As String
Dim Posizione_orbitale1 As String
Dim Canale1 As String
Dim System1 As String
Dim Frequenza1 As String
Dim Polarizzazione1 As String
Dim Bit_Rate1 As String
Dim Symbol_Rate1 As String
Dim MTT1 As String
Dim FEC1 As String
Dim Origine1 As String
Dim Destinazione1 As String
Dim Miaf1 As String
Dim Mezzo_ricez1 As String
Dim Banda1 As String
Dim satellite2 As String
Dim Cod_satellite2 As String
Dim NRO_Cod_satellite2 As String
Dim Posizione_orbitale2 As String
Dim Canale2 As String
Dim System2 As String
Dim Frequenza2 As String
Dim Polarizzazione2 As String
Dim Bit_Rate2 As String
Dim Symbol_Rate2 As String
Dim MTT2 As String
Dim FEC2 As String
Dim Origine2 As String
Dim Destinazione2 As String
Dim Miaf2 As String
Dim Mezzo_ricez2 As String
Dim Banda2 As String
' M.G. Nuovi CodiceServizio e attributi satellitari
Dim CM_Freq_Dw_link As String
Dim CM_Pol_Dw_link As String
Dim CM_Codice_Servizio As String
Dim C2_Freq_Dw_link As String
Dim C2_Pol_Dw_link As String
Dim C2_Codice_Servizio As String
' M.G. -----------------------------------------



CM_Ora_Iniz_PM = "1900/01/01 00:00.00"
CM_Ora_Fine_PM = "1900/01/01 00:00.00"
satellite1 = ""
Cod_satellite1 = ""
NRO_Cod_satellite1 = ""
Posizione_orbitale1 = ""
Canale1 = ""
System1 = ""
Frequenza1 = ""
Polarizzazione1 = ""
Bit_Rate1 = ""
Symbol_Rate1 = ""
MTT1 = ""
FEC1 = ""
Origine1 = ""
Destinazione1 = ""
Miaf1 = ""
Mezzo_ricez1 = ""
Banda1 = ""
satellite2 = ""
Cod_satellite2 = ""
NRO_Cod_satellite2 = ""
Posizione_orbitale2 = ""
Canale2 = ""
System2 = ""
Frequenza2 = ""
Polarizzazione2 = ""
Bit_Rate2 = ""
Symbol_Rate2 = ""
MTT2 = ""
FEC2 = ""
Origine2 = ""
Destinazione2 = ""
Miaf2 = ""
Mezzo_ricez2 = ""
Banda2 = ""
sist1 = ""
tel1 = ""
loc1 = ""
txp1 = ""
audio1_1 = ""
audio2_1 = ""
note1 = ""
C2_Ora_Iniz_PM = "1900/01/01 00:00.00"
C2_Ora_Fine_PM = "1900/01/01 00:00.00"
sist2 = ""
tel2 = ""
loc2 = ""
txp2 = ""
audio1_2 = ""
audio2_2 = ""
note2 = ""

CM_Freq_Dw_link = ""
CM_Pol_Dw_link = ""
CM_Codice_Servizio = ""
C2_Freq_Dw_link = ""
C2_Pol_Dw_link = ""
C2_Codice_Servizio = ""

'VERIFICA SE C'E' GIA' STATO UN RIAGGANCIO  (no collegamenti nella tabella confermebooking ma collegamenti mobili booking gia' inseriti per quella versione)
sql = "select * from CollegamentiMobili Where " & _
      "anno_rich_cm=" & anno & " and num_rich_cm=" & num & " and ver_rich_cm=" & ver
Set rs = cn.execute(sql)
If Not rs.EOF Then
    Riaggancia_Versione_prec = True
    Exit Function
End If
rs.Close

'verifica se ci sono CM nella versione precedente della richiesta
ult_ver_cm = getVersione("CM", cn, anno, num, ver_prec)

If ult_ver_cm = 0 Then
WriteLog "non trovata versione cm"
    Exit Function
Else
    sql = "select * from CollegamentiMobili Where " & _
          "anno_rich_cm=" & anno & " and num_rich_cm=" & num & " and ver_rich_cm=" & ver_prec & " and versione_cm = " & ult_ver_cm
    Set rs = cn.execute(sql)
    If rs.EOF Then
WriteLog "non ci sono collegamenti mobili precedenti"
        Exit Function
    Else
        'INSERIMENTO COLLEGAMENTI MOBILI DA VERSIONE PRECEDENTE DELLA RICHIESTA (VERSIONE_cm + RECENTE)
        '-------------------------------------------------------------------------------------------------------------------------
        If IsNull(rs("C1_Ora_Inizio_CM")) Then
             CM_Ora_Iniz_PM = "1900/01/01 00:00:00"
        Else
             CM_Ora_Iniz_PM = "1900/01/01 " & Right("00" & Hour(rs("C1_Ora_Inizio_CM")), 2) & ":" & Right("00" & Minute(rs("C1_Ora_Inizio_CM")), 2) & ":" & Right("00" & Second(rs("C1_Ora_Inizio_CM")), 2)
        End If
        If IsNull(rs("C1_Ora_Fine_CM")) Then
             CM_Ora_Fine_PM = "1900/01/01 00:00:00"
        Else
             CM_Ora_Fine_PM = "1900/01/01 " & Right("00" & Hour(rs("C1_Ora_Fine_CM")), 2) & ":" & Right("00" & Minute(rs("C1_Ora_Fine_CM")), 2) & ":" & Right("00" & Second(rs("C1_Ora_Fine_CM")), 2)
        End If
    
        sist1 = Replace(rs("C1_Sistema_CM") & "", "'", "''")
        tel1 = Replace(rs("C1_Telefono_CM") & "", "'", "''")
        loc1 = Replace(rs("C1_Localita_RX_CM") & "", "'", "''")
        txp1 = Replace(rs("C1_Nome_TXP_CM") & "", "'", "''")
        audio1_1 = Replace(rs("C1_Audio1_CM") & "", "'", "''")
        audio2_1 = Replace(rs("C1_Audio2_CM") & "", "'", "''")
        note1 = Replace(rs("C1_Note_CM") & "", "'", "''")
        satellite1 = Replace(rs("CM_Satellite") & "", "'", "''")
        NRO_Cod_satellite1 = Replace(rs("CM_NRO_Cod_satellite") & "", "'", "''")
        Cod_satellite1 = Replace(rs("CM_Cod_satellite") & "", "'", "''")
        Posizione_orbitale1 = Replace(rs("CM_Posizione_orbitale") & "", "'", "''")
        Canale1 = Replace(rs("CM_Canale") & "", "'", "''")
        System1 = Replace(rs("CM_SYSTEM") & "", "'", "''")
        Frequenza1 = Replace(rs("CM_Frequenza") & "", "'", "''")
        Polarizzazione1 = Replace(rs("CM_Polarizzazione") & "", "'", "''")
        Bit_Rate1 = Replace(rs("CM_Bit_Rate") & "", "'", "''")
        Symbol_Rate1 = Replace(rs("CM_Symbol_Rate") & "", "'", "''")
        MTT1 = Replace(rs("CM_MTT") & "", "'", "''")
        FEC1 = Replace(rs("CM_FEC") & "", "'", "''")
        Origine1 = Replace(rs("CM_Origine") & "", "'", "''")
        Destinazione1 = Replace(rs("CM_Destinazione") & "", "'", "''")
        Miaf1 = Replace(rs("C1_Miaf_CM") & "", "'", "''")
        Mezzo_ricez1 = Replace(rs("C1_Mezzo_ricez_CM") & "", "'", "''")
        Banda1 = Replace(rs("CM_Banda") & "", "'", "''")
        
        CM_Freq_Dw_link = Replace(rs("CM_Freq_Dw_link") & "" & "", "'", "''")
        CM_Pol_Dw_link = Replace(rs("CM_Pol_Dw_link") & "" & "", "'", "''")
        CM_Codice_Servizio = Replace(rs("CM_Codice_Servizio") & "" & "", "'", "''")
        
        
        If IsNull(rs("C2_Ora_Inizio_CM")) Then
             C2_Ora_Iniz_PM = "1900/01/01 00:00:00"
        Else
             C2_Ora_Iniz_PM = "1900/01/01 " & Right("00" & Hour(rs("C2_Ora_Inizio_CM")), 2) & ":" & Right("00" & Minute(rs("C2_Ora_Inizio_CM")), 2) & ":" & Right("00" & Second(rs("C2_Ora_Inizio_CM")), 2)
        End If
        If IsNull(rs("C2_Ora_Fine_CM")) Then
             C2_Ora_Fine_PM = "1900/01/01 00:00:00"
        Else
             C2_Ora_Fine_PM = "1900/01/01 " & Right("00" & Hour(rs("C2_Ora_Fine_CM")), 2) & ":" & Right("00" & Minute(rs("C2_Ora_Fine_CM")), 2) & ":" & Right("00" & Second(rs("C2_Ora_Fine_CM")), 2)
        End If
        sist2 = Replace(rs("C2_Sistema_CM") & "", "'", "''")
        tel2 = Replace(rs("C2_Telefono_CM") & "", "'", "''")
        loc2 = Replace(rs("C2_Localita_RX_CM") & "", "'", "''")
        txp2 = Replace(rs("C2_Nome_TXP_CM") & "", "'", "''")
        audio1_2 = Replace(rs("C2_Audio1_CM") & "", "'", "''")
        audio2_2 = Replace(rs("C2_Audio2_CM") & "", "'", "''")
        note2 = Replace(rs("C2_Note_CM") & "", "'", "''")
        satellite2 = Replace(rs("C2_Satellite") & "", "'", "''")
        NRO_Cod_satellite2 = Replace(rs("C2_NRO_Cod_satellite") & "", "'", "''")
        Cod_satellite2 = Replace(rs("C2_Cod_satellite") & "", "'", "''")
        Posizione_orbitale2 = Replace(rs("C2_Posizione_orbitale") & "", "'", "''")
        Canale2 = Replace(rs("C2_Canale") & "", "'", "''")
        System2 = Replace(rs("C2_SYSTEM") & "", "'", "''")
        Frequenza2 = Replace(rs("C2_Frequenza") & "", "'", "''")
        Polarizzazione2 = Replace(rs("C2_Polarizzazione") & "", "'", "''")
        Bit_Rate2 = Replace(rs("C2_Bit_Rate") & "", "'", "''")
        Symbol_Rate2 = Replace(rs("C2_Symbol_Rate") & "", "'", "''")
        MTT2 = Replace(rs("C2_MTT") & "", "'", "''")
        FEC2 = Replace(rs("C2_FEC") & "", "'", "''")
        Origine2 = Replace(rs("C2_Origine") & "", "'", "''")
        Destinazione2 = Replace(rs("C2_Destinazione") & "", "'", "''")
        Miaf2 = Replace(rs("C2_Miaf_CM") & "", "'", "''")
        Mezzo_ricez2 = Replace(rs("C2_Mezzo_ricez_CM") & "", "'", "''")
        Banda2 = Replace(rs("C2_Banda") & "", "'", "''")
        
        C2_Freq_Dw_link = Replace(rs("C2_Freq_Dw_link") & "", "'", "''")
        C2_Pol_Dw_link = Replace(rs("C2_Pol_Dw_link") & "", "'", "''")
        C2_Codice_Servizio = Replace(rs("C2_Codice_Servizio") & "", "'", "''")
        
   
        sql = "INSERT INTO CollegamentiMobili (" & _
        "Anno_Rich_CM, Num_Rich_CM, Ver_Rich_CM, Versione_CM," & _
        "CM_satellite, CM_Cod_satellite, CM_NRO_Cod_satellite ,CM_Posizione_orbitale, " & _
        "CM_Canale, CM_System, CM_Frequenza, CM_Polarizzazione, " & _
        "CM_Bit_Rate, CM_Symbol_Rate, CM_MTT , CM_FEC, " & _
        "CM_Origine, CM_Destinazione, C1_Miaf_CM, C1_Mezzo_ricez_CM , CM_Banda , " & _
        "C2_satellite, C2_Cod_satellite, C2_NRO_Cod_satellite ,C2_Posizione_orbitale, " & _
        "C2_Canale, C2_System, C2_Frequenza, C2_Polarizzazione, " & _
        "C2_Bit_Rate, C2_Symbol_Rate, C2_MTT , C2_FEC, " & _
        "C2_Origine, C2_Destinazione, C2_Miaf_CM, C2_Mezzo_ricez_CM , C2_Banda , " & _
        "C1_Sistema_CM, C1_Telefono_CM, C1_Nome_TXP_CM, C1_Localita_RX_CM," & "C1_Audio1_CM, C1_Audio2_CM, C1_Ora_Inizio_CM, C1_Ora_Fine_CM, C1_Note_CM," & "C2_Sistema_CM, C2_Telefono_CM, C2_Nome_TXP_CM, C2_Localita_RX_CM," & "C2_Audio1_CM, C2_Audio2_CM, C2_Ora_Inizio_CM, C2_Ora_Fine_CM, C2_Note_CM, CM_Freq_Dw_link, CM_Pol_Dw_link, CM_Codice_Servizio, C2_Freq_Dw_link, C2_Pol_Dw_link, C2_Codice_Servizio" & _
        ") VALUES (" & _
            anno & ", " & num & ", " & ver & ", " & 1 & " , " & _
            "'" & satellite1 & "', '" & Cod_satellite1 & "', " & "'" & NRO_Cod_satellite1 & "', '" & Posizione_orbitale1 & "', " & _
            "'" & Canale1 & "', '" & System1 & "', " & "'" & Frequenza1 & "', '" & Polarizzazione1 & "', " & _
            "'" & Bit_Rate1 & "', '" & Symbol_Rate1 & "', " & "'" & MTT1 & "', '" & FEC1 & "', " & _
            "'" & Origine1 & "', '" & Destinazione1 & "', " & "'" & Miaf1 & "', '" & Mezzo_ricez1 & "', '" & Banda1 & "', " & _
            "'" & satellite2 & "', '" & Cod_satellite2 & "', " & "'" & NRO_Cod_satellite2 & "', '" & Posizione_orbitale2 & "', " & _
            "'" & Canale2 & "', '" & System2 & "', " & "'" & Frequenza2 & "', '" & Polarizzazione2 & "', " & _
            "'" & Bit_Rate2 & "', '" & Symbol_Rate2 & "', " & "'" & MTT2 & "', '" & FEC2 & "', " & _
            "'" & Origine2 & "', '" & Destinazione2 & "', " & "'" & Miaf2 & "', '" & Mezzo_ricez2 & "', '" & Banda2 & "', " & _
            "'" & sist1 & "', '" & tel1 & "', " & "'" & txp1 & "', '" & loc1 & "', " & "'" & audio1_1 & "', '" & audio2_1 & "', '" & CM_Ora_Iniz_PM & "', '" & CM_Ora_Fine_PM & "', " & "'" & note1 & "', " & "'" & sist2 & "', '" & tel2 & "', " & _
            "'" & txp2 & "', '" & loc2 & "', " & "'" & audio1_2 & "', '" & audio2_2 & "', '" & C2_Ora_Iniz_PM & "', '" & C2_Ora_Fine_PM & "', '" & note2 & "', '" & CM_Freq_Dw_link & "', '" & CM_Pol_Dw_link & "', '" & CM_Codice_Servizio & "', '" & C2_Freq_Dw_link & "', '" & C2_Pol_Dw_link & "', '" & C2_Codice_Servizio & "')"
        
        '-------------------------------------------------------------------------------------------------------------------------
        'FINE INSERIMENTO
    cn.execute sql
    End If
End If

Riaggancia_Versione_prec = True
Exit Function
err_riaggancio:
    WriteLog "Errore nella funzione di riaggancio versione precedente CM: " & Err.Description & vbCrLf & "SQL : " & sql
End Function

Private Function aggiornaFlagOttimConFono(cn As Object, ByVal anno As Integer, ByVal num As Long) As Boolean
On Error GoTo err_aggiornaFlagOttimConFono
Dim sql, ret

    sql = "update ottim_fono set flag_variante ='S' where num_ottim=" & num & " and anno_ottim=" & anno
    cn.execute sql, ret
    If ret > 0 Then
        WriteLog "Aggiornato flag fono della ottim. " & anno & "/" & num
    End If
    aggiornaFlagOttimConFono = True
    Exit Function

err_aggiornaFlagOttimConFono:
    WriteLog "Errore nella funzione aggiornaFlagOttimConFono: " & Err.Description & vbCrLf & "SQL : " & sql
    aggiornaFlagOttimConFono = False
End Function

Private Function risorsaConOttimizzazioni(cn As Object, ByVal anno As Integer, ByVal num As Long, ByVal ver As Integer) As Boolean
Dim sql
Dim rs As Object

On Error GoTo err_risorsaConOttimizzazioni
    
    risorsaConOttimizzazioni = True
    sql = "SELECT DISTINCT (num_ottim), anno_ottim, stato_risorsa, id_ver_rich " & _
    " FROM Risorse WHERE ID_Anno_Rich=" & anno & " and " & _
    " ID_Num_Rich=" & num & " and ID_Ver_Rich <= " & ver & " and not(anno_ottim is null)"
    
    Set rs = CreateObject("ADODB.Recordset")
    Set rs.ActiveConnection = cn
    rs.CursorType = 3 'adOpenStatic
    rs.Open sql
    
    Do While (Not rs.EOF)
        If Not aggiornaFlagOttimConFono(cn, rs("anno_ottim"), rs("num_ottim")) Then
            risorsaConOttimizzazioni = False
            Exit Do
        End If
        rs.movenext
    Loop
    rs.Close
    Set rs = Nothing
    
    Exit Function

err_risorsaConOttimizzazioni:
    WriteLog "Errore nella funzione risorsaConOttimizzazioni: " & Err.Description & vbCrLf & "SQL : " & sql
    Set rs = Nothing
    risorsaConOttimizzazioni = False
End Function

