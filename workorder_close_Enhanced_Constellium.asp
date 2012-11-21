<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>
<!--#INCLUDE FILE="../common/mc_all.asp" -->
<%
Response.AddHeader "cache-control","no-store"

Dim db, rs, rsa, rsLAType, rsLocation, rsAccounts, rsCategory, rsLA, rsINV, rsVEN, rsTasks, rsTaskCnt, pmrs, pmsql, sql, errormessage
Dim ArrayLA, ArrayIN, ArrayLAT, ArrayLOC, ArrayAC, ArrayCA, ArrayVC
Dim WOPK, WOID, Assetpk, Mode, curmod, DialogTitle, RightTitle, TaskDefaultInitials
Dim TZ
Dim requested, requesteddate, requestedtime, requestedinitials
Dim issued, issueddate, issuedtime, issuedinitials
Dim responded, respondeddate, respondedtime, respondedinitials
Dim completed, completeddate, completedtime, completedinitials
Dim finalized, finalizeddate, finalizedtime, finalizedinitials
Dim closed, closeddate, closedtime, closedinitials
Dim prefvalue, prefdesc, prefpk, RCPK
Dim WO_CLOSE_STATUSDATESREADONLY
Dim WO_CLOSE_SHOWLABORREPORT_REQ
Dim WO_CLOSE_SHOWMETERREADINGS_REQM1
Dim WO_CLOSE_SHOWMETERREADINGS_REQM2
Dim WO_CLOSE_SHOWACCOUNTCATEGORY_AREQ
Dim WO_CLOSE_SHOWACCOUNTCATEGORY_CREQ
Dim WO_CLOSE_SHOWFAILUREANALYSIS_PREQ
Dim WO_CLOSE_SHOWFAILUREANALYSIS_FREQ
Dim WO_CLOSE_SHOWFAILUREANALYSIS_SREQ
Dim WO_CLOSE_SHOWLABORACTUAL_REQ
Dim WO_CLOSE_SHOWPARTACTUAL_REQ
Dim WO_CLOSE_SHOWMISCCOSTACTUAL_REQ
Dim WO_CLOSE_ALLTASKSCOMPLETE_REQ
Dim WO_CLOSE_REQ_ASSIGNMENTS_FOR_CLOSE
Dim WO_CLOSE_SPLITLABORHOURS_CBDEFAULT
Dim WO_CLOSE_CHECKMETERREADINGDELTA
Dim WO_CLOSE_UPDATEASSETMETERS
Dim WO_CLOSE_SHOWLABORACTUAL_FILTERRC 
Dim WO_CLOSE_SHOWONDEMAND_FOLLOWUP_REQ
'Dim WO_CLOSE_LABORREPORTOPTIONS
Dim laborreport
Dim lri
Dim txtwogrouptype
Dim txtaccountpk
Dim txtaccount
Dim txtaccountdesc
Dim txtchargeable
Dim txtAccountAll
Dim txtcategory
Dim txtcategorydesc
Dim txtcategorypk
Dim txtCategoryAll
Dim txtwo
Dim txtreason
Dim txtassetpk
Dim txtTasks
Dim txtTaskInitials
Dim txtLabor1
Dim txtLabor3
Dim txtMyLaborHrs
Dim txtMaterials
Dim txtOtherCost
Dim txtproblempk
Dim txtproblem
Dim txtproblemdesc
Dim txtfailurepk
Dim txtfailure
Dim txtfailuredesc
Dim txtsolutionpk
Dim txtsolution
Dim txtsolutiondesc
Dim txtfailurewo
Dim txtFollowUpAllWO
Dim txtFollowUpSingleWO
Dim txtFollowUpMultiWO
Dim txtmeter1reading
Dim txtmeter2reading
Dim txtisup
Dim txtAssetStatusHistoryPK
Dim actionwhere
Dim assetexists
Dim ismeter
Dim AssetPhoto
Dim bgposition
Dim WOGroupPK
Dim txtWOGroupAll
Dim wogroupischecked
Dim txtDrawingUpdatesNeeded
Dim getassignments
Dim MeterPMs
Dim Meter1PMIntervals
Dim Meter1PMNext
Dim Meter2PMIntervals
Dim Meter2PMNext
Dim Meter1,Meter2
Dim assetmeter1, assetmeter2
Dim lpk, lid, lnm, lrcpk
Dim showLaborActual, showPartActual, showMiscCostActual
Dim AllTasksComplete, TaskMessage, TaskCode
AllTasksComplete="No"
TaskCode=0

assetmeter1=0
assetmeter2=0
MeterPMs=""
Meter1PMIntervals=""
Meter1PMNext=""
Meter2PMIntervals=""
Meter2PMNext=""

CurMod = Trim(Request("CurMod"))
Mode = Trim(Request("Mode"))
WOPK = Trim(Request("WOPK"))
txtwogrouptype = ""
WOGroupPK = Trim(Request("WOGroupPK"))
If WOGroupPK = "" Then
	WOGroupPK = "-1"
End If
actionwhere = Trim(Request("actionwhere"))
assetphoto = ""
bgposition = "97% 88%"
wogroupischecked = UCase(Trim(Request.QueryString("wogroupischecked")))
If wogroupischecked = "Y" Then
	wogroupischecked = True
Else
	wogroupischecked = False
End If
getassignments = False

TaskDefaultInitials = GetSession("UserInitials")

'Response.Write actionwhere
'Response.End

Set db = New ADOHelper
errormessage = ""

' Added 12/13/2004
Dim AccessToRespond, AccessToComplete, AccessToFinalize, AccessToClose, AccessToStatusDateTime

AccessToRespond = GetAccessRight(db,"WORespond",0)
AccessToComplete = GetAccessRight(db,"WOComplete",0)
AccessToFinalize = GetAccessRight(db,"WOFinalize",0)
AccessToClose = GetAccessRight(db,"WOCloseFromDialog",0)
AccessToStatusDateTime = GetAccessRight(db,"WOStatusDateTime",0)

'RGJ BEGIN
If Not GetSession("DFF") = "" Then
  If GetSession("DFF") = "Y" Then
    TZ = "Y"
  Else
    TZ = "N"
  End If
End If
'Get User Information
'Get LaborPK
If Not GetSession("UserPK") = "" Then
  lpk=GetSession("UserPK")
  lid=GetSession("UserID")
  lnm=GetSession("UserName")
Else
    lpk=""
    lid=""
    lnm=""
End If

'RGJ BEGIN
Call GetPreference(db,False,"WO_CLOSE_STATUSDATESREADONLY",prefvalue, prefdesc, prefpk)
WO_CLOSE_STATUSDATESREADONLY = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWLABORREPORT_REQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWLABORREPORT_REQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWMETERREADINGS_REQM1",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWMETERREADINGS_REQM1 = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWMETERREADINGS_REQM2",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWMETERREADINGS_REQM2 = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWACCOUNTCATEGORY_AREQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWACCOUNTCATEGORY_AREQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWACCOUNTCATEGORY_CREQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWACCOUNTCATEGORY_CREQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWFAILUREANALYSIS_PREQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWFAILUREANALYSIS_PREQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWFAILUREANALYSIS_FREQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWFAILUREANALYSIS_FREQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWFAILUREANALYSIS_SREQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWFAILUREANALYSIS_SREQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWLABORACTUAL_REQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWLABORACTUAL_REQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWPARTACTUAL_REQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWPARTACTUAL_REQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_SHOWMISCCOSTACTUAL_REQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWMISCCOSTACTUAL_REQ = prefvalue
Call GetPreference(db,False,"WO_CLOSE_REQ_ASSIGNMENTS_FOR_CLOSE",prefvalue, prefdesc, prefpk)
WO_CLOSE_REQ_ASSIGNMENTS_FOR_CLOSE = prefvalue

'Added to bypass meter checking - April 2010
Call GetPreference(db,False,"WO_CLOSE_CHECKMETERREADINGDELTA",prefvalue, prefdesc, prefpk)
WO_CLOSE_CHECKMETERREADINGDELTA = prefvalue
If NullCheck(WO_CLOSE_CHECKMETERREADINGDELTA) = "" Then
  WO_CLOSE_CHECKMETERREADINGDELTA = "Yes"
End If
Call GetPreference(db,False,"WO_CLOSE_UPDATEASSETMETERS",prefvalue, prefdesc, prefpk)
WO_CLOSE_UPDATEASSETMETERS = prefvalue
If NullCheck(WO_CLOSE_UPDATEASSETMETERS) = "" Then
  WO_CLOSE_UPDATEASSETMETERS = "Yes"
End If
'Added to filter Labor to RC of logged in user - 5/10/2010
Call GetPreference(db,False,"WO_CLOSE_SHOWLABORACTUAL_FILTERRC",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWLABORACTUAL_FILTERRC = prefvalue
'Default to NO if pref DNE
If NullCheck(WO_CLOSE_SHOWLABORACTUAL_FILTERRC) = "" Then
  WO_CLOSE_SHOWLABORACTUAL_FILTERRC = "No"
End If
'Added to not load if item is not selected
Call GetPreference(db,False,"WO_CLOSE_SHOWLABORACTUAL",prefvalue, prefdesc, prefpk)
showLaborActual = prefvalue
If NullCheck(showLaborActual) = "" Then
  showLaborActual = "Yes"
End If
Call GetPreference(db,False,"WO_CLOSE_SHOWPARTACTUAL",prefvalue, prefdesc, prefpk)
showPartActual = prefvalue
If NullCheck(showPartActual) = "" Then
  showPartActual = "Yes"
End If
Call GetPreference(db,False,"WO_CLOSE_SHOWMISCCOSTACTUAL",prefvalue, prefdesc, prefpk)
showMiscCostActual = prefvalue
If NullCheck(showMiscCostActual) = "" Then
  showMiscCostActual = "Yes"
End If
'RGJ V5 Addition - Require All Tasks to be completed before allowing closure of WO
Call GetPreference(db,False,"WO_CLOSE_ALLTASKSCOMPLETE_REQ",prefvalue, prefdesc, prefpk)
WO_CLOSE_ALLTASKSCOMPLETE_REQ = prefvalue
'Default to NO if pref DNE
If NullCheck(WO_CLOSE_ALLTASKSCOMPLETE_REQ) = "" Then
  WO_CLOSE_ALLTASKSCOMPLETE_REQ = "No"
End If

'RGJ V5 Addition - Default Split Labor Hours to checked or unchecked
Call GetPreference(db,False,"WO_CLOSE_SPLITLABORHOURS_CBDEFAULT",prefvalue, prefdesc, prefpk)
WO_CLOSE_SPLITLABORHOURS_CBDEFAULT = prefvalue
'Default to NO if pref DNE
If NullCheck(WO_CLOSE_SPLITLABORHOURS_CBDEFAULT) = "" Then
  WO_CLOSE_SPLITLABORHOURS_CBDEFAULT = "No"
End If

Call GetPreference(db,False,"WO_CLOSE_SHOWONDEMAND_FOLLOWUP",prefvalue, prefdesc, prefpk)
WO_CLOSE_SHOWONDEMAND_FOLLOWUP_REQ = prefvalue
'Default to NO if pref DNE
If NullCheck(WO_CLOSE_SHOWONDEMAND_FOLLOWUP_REQ) = "" Then
  WO_CLOSE_SHOWONDEMAND_FOLLOWUP_REQ = "No"
End If

'Response.Write "WO_CLOSE_SHOWFAILUREANALYSIS_PREQ: " & WO_CLOSE_SHOWFAILUREANALYSIS_PREQ & "<br>"
'Response.Write "WO_CLOSE_SHOWFAILUREANALYSIS_FREQ: " & WO_CLOSE_SHOWFAILUREANALYSIS_FREQ & "<br>"
'Response.Write "WO_CLOSE_SHOWFAILUREANALYSIS_SREQ: " & WO_CLOSE_SHOWFAILUREANALYSIS_SREQ & "<br>"

'RGJ Labor Report Enhancedment V5- SP1 --> WO_CLOSE_LABORREPORTOPTIONS
'Call GetPreference(db,False,"WO_CLOSE_LABORREPORTOPTIONS",prefvalue, prefdesc, prefpk)
'WO_CLOSE_LABORREPORTOPTIONS = prefvalue
'Default to NO if pref DNE
'If NullCheck(WO_CLOSE_LABORREPORTOPTIONS) = "" Then
'  WO_CLOSE_LABORREPORTOPTIONS = "1"
'End If

'Get Repair Center access
If WO_CLOSE_SHOWLABORACTUAL_FILTERRC = "Yes" Then
  If NullCheck(GetSession("RCDENY")) <> "" Then
    lrcpk = "SELECT RepairCenterPK FROM RepairCenter WHERE Active = 1 AND RepairCenterPK IN (" & GetSession("RCDENY") & ")"
  Else
    lrcpk = "SELECT RepairCenterPK FROM RepairCenter WHERE Active = 1 AND RepairCenterPK IN (" & GetSession("RCPK") & ")"
  End If
Else
  lrcpk = ""
End If
'RGJ END

Call GetData

'AllTasksComplete
If TaskCode = -1 Then
  AllTasksComplete="NO"
Else ' TaskCode = 1 or 0
  AllTasksComplete="YES"
End If

checkforsubmit

Set db = New ADOHelper

If mode = "WO" Then

	sql = _
	"SELECT WO.*, Asset.ISUP, Asset.IsMeter, Asset.Meter1Reading As A_Meter1Reading, Asset.Meter2Reading As A_Meter2Reading, REPLACE(Asset.Photo,'_MAIN','_WO') AS AssetPhoto " & _
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON ASSET.ASSETPK = WO.ASSETPK " & _
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")

	If Not db.dok Then
		errormessage = "There was a problem accessing the Work Order record. Please contact your maintenance manager for support.<br><br>" & db.derror
	Else
		If rs.eof Then
			errormessage = "There was a problem accessing the Work Order record (not found). Please contact your maintenance manager for support."
		Else
		  assetmeter1 = Trim(rs("A_Meter1Reading"))
		  assetmeter2 = Trim(rs("A_Meter2Reading"))

			If rs("IsOpen") Then
				woid = NullCheck(rs("woid"))
				txtwogrouptype = NullCheck(rs("wogrouptype"))
				If Not NullCheck(rs("wogrouppk")) = "" Then
					WOGroupPK = NullCheck(rs("wogrouppk"))
				End If
				assetpk = NullCheck(rs("assetpk"))
                RCPK = Nullcheck(rs("RepairCenterPK"))
				requesteddate = DateNullCheck(rs("requested"))
				requestedtime = TimeNullCheckAT(rs("requested"))
				requestedinitials = NullCheck(rs("takenbyinitials"))
				If requesteddate = "" Then
					requesteddate = DateNullCheck(Date())
					requestedtime = FixTime(TimeNullCheckAT(Time()))
					requested = False
				Else
					requested = True
				End If

				issueddate = DateNullCheck(rs("issued"))
				issuedtime = TimeNullCheckAT(rs("issued"))
				issuedinitials = NullCheck(rs("issuedinitials"))
				If issueddate = "" Then
					issueddate = DateNullCheck(Date())
					issuedtime = FixTime(TimeNullCheckAT(Time()))
					issued = False
				Else
					issued = True
				End If
				If UCase(rs("status")) = "REQUESTED" and issuedinitials = "" Then
					issuedinitials = GetSession("UserInitials")
				End If

				respondeddate = DateNullCheck(rs("responded"))
				respondedtime = TimeNullCheckAT(rs("responded"))
				respondedinitials = NullCheck(rs("respondedinitials"))
				If respondeddate = "" Then
					respondeddate = DateNullCheck(Date())
					respondedtime = FixTime(TimeNullCheckAT(Time()))
					responded = False
				Else
					responded = True
				End If

				completeddate = DateNullCheck(rs("complete"))
				completedtime = TimeNullCheckAT(rs("complete"))
				completedinitials = NullCheck(rs("completedinitials"))
				If completeddate = "" Then
					completeddate = DateNullCheck(Date())
					completedtime = FixTime(TimeNullCheckAT(Time()))
					completed = False
				Else
					completed = True
				End If

				finalizeddate = DateNullCheck(rs("finalized"))
				finalizedtime = TimeNullCheckAT(rs("finalized"))
				finalizedinitials = NullCheck(rs("finalizedinitials"))
				If finalizeddate = "" Then
					finalizeddate = DateNullCheck(Date())
					finalizedtime = FixTime(TimeNullCheckAT(Time()))
					finalized = False
				Else
					finalized = True
				End If

				closeddate = DateNullCheck(rs("closed"))
				closedtime = TimeNullCheckAT(rs("closed"))
				closedinitials = NullCheck(rs("closedinitials"))
				If closeddate = "" Then
					closeddate = DateNullCheck(Date())
					closedtime = FixTime(TimeNullCheckAT(Time()))
					closed = False
				Else
					closed = True
				End If

				laborreport = Replace(NullCheck(rs("laborreport")),"%0D%0A",CHR(13) & CHR(10))
				txtaccountpk = NullCheck(rs("accountpk"))
				txtaccount = NullCheck(rs("accountid"))
				txtaccountdesc = NullCheck(rs("accountname"))
				txtcategorypk = NullCheck(rs("categorypk"))
				txtcategory = NullCheck(rs("categoryid"))
				txtcategorydesc = NullCheck(rs("categoryname"))
				txtchargeable = rs("chargeable")
				txtproblempk = NullCheck(rs("problempk"))
				txtproblem = NullCheck(rs("problemid"))
				txtproblemdesc = NullCheck(rs("problemname"))
				txtfailurepk = NullCheck(rs("failurepk"))
				txtfailure = NullCheck(rs("failureid"))
				txtfailuredesc = NullCheck(rs("failurename"))
				txtsolutionpk = NullCheck(rs("solutionpk"))
				txtsolution = NullCheck(rs("solutionid"))
				txtsolutiondesc = NullCheck(rs("solutionname"))
				txtfailurewo = rs("failedwo")
				txtreason = NullCheck(rs("reason"))
				txtwo = NullCheck(rs("woid"))
				txtassetpk = NullCheck(rs("assetpk"))
				If NullCheck(rs("meter1reading")) = "" or NullCheck(rs("meter1reading")) = "0" Then
					txtmeter1reading = NullCheck(rs("a_meter1reading"))
				Else
					txtmeter1reading = NullCheck(rs("meter1reading"))
				End If
				If NullCheck(rs("meter2reading")) = "" or NullCheck(rs("meter2reading")) = "0" Then
					txtmeter2reading = NullCheck(rs("a_meter2reading"))
				Else
					txtmeter2reading = NullCheck(rs("meter2reading"))
				End If
				If Not IsNull(rs("IsUp")) Then
					txtisup = rs("isup")
					ismeter = rs("ismeter")
					assetexists = True
				Else
					txtisup = True
					ismeter = False
					assetexists = False
				End If
				txtAssetStatusHistoryPK = NullCheck(rs("AssetStatusHistoryPK"))
				If Not NullCheck(rs("assetphoto")) = "" Then
					assetphoto = getclientimage(Trim(rs("assetphoto")))
					bgposition = "55% 83%"
				End If
				getassignments = True
			Else
				errormessage = "The selected Work Order is already either Closed, Canceled, or Denied."
			End If
		End If
	End If

  'RGJ BEGIN
	pmsql = "SELECT PM.PMPK, Asset.Meter1Reading, Asset.Meter2Reading, Meter1Interval, Meter1NextInterval, Meter2Interval, Meter2NextInterval " & vbCrLf &_
  "FROM WO WITH (NOLOCK) INNER JOIN Asset ON Asset.AssetPK = WO.AssetPK INNER JOIN PMAsset ON PMAsset.AssetPK = Asset.AssetPK INNER JOIN PM ON PM.PMPK = PMAsset.PMPK  " & vbCrLf &_
  "WHERE Frequency = 'METER' AND WO.WOPK = " & WOPK & " "
  'Response.write "<textarea rows=4 cols=100>" & pmsql & "</textarea>"
  'Response.End
	
  Set pmrs = db.RunSQLReturnRS(pmsql,"")
	If Not db.dok Then
		errormessage = "There was a problem accessing the Work Order record. Please contact your maintenance manager for support.<br><br>" & db.derror
	Else
		If Not pmrs.eof Then
      Do While Not pmrs.EOF
        'Fill Variables with data
        MeterPMs = MeterPMs + CStr(pmrs("PMPK")) + ","
        Meter1PMIntervals = Meter1PMIntervals + CStr(pmrs("Meter1Interval")) + ","
        Meter1PMNext = Meter1PMNext + CStr(Trim(pmrs("Meter1NextInterval"))) + ","
        Meter2PMIntervals = Meter2PMIntervals + CStr(pmrs("Meter2Interval")) + ","
        Meter2PMNext = Meter2PMNext + CStr(Trim(pmrs("Meter2NextInterval"))) + ","

        pmrs.MoveNext
      Loop
    Else
      'Set Vars to blank string
      MeterPMs=""
      Meter1PMIntervals=""
      Meter1PMNext=""
      Meter2PMIntervals=""
      Meter2PMNext=""
    End If
 	End If
  Call CloseObj(pmrs)

  'Remove comma at end if it exists
  If InStr(1,MeterPMs,",") Then
    MeterPMs = Trim(Mid(MeterPMs,1,Len(MeterPMs)-1))
  End If
  If InStr(1,Meter1PMIntervals,",") Then
    Meter1PMIntervals = Trim(Mid(Meter1PMIntervals,1,Len(Meter1PMIntervals)-1))
  End If
  If InStr(1,Meter1PMNext,",") Then
    Meter1PMNext = Trim(Mid(Meter1PMNext,1,Len(Meter1PMNext)-1))
  End If
  If InStr(1,Meter2PMIntervals,",") Then
    Meter2PMIntervals = Trim(Mid(Meter2PMIntervals,1,Len(Meter2PMIntervals)-1))
  End If
  If InStr(1,Meter2PMNext,",") Then
    Meter2PMNext = Trim(Mid(Meter2PMNext,1,Len(Meter2PMNext)-1))
  End If
  'RGJ END


Else 'If mode = "WO" Then

	woid = ""
	txtwogrouptype = ""
	assetpk = ""
	RCPK = 1
	requested = False
	issued = False
	responded = False
	completed = False
	finalized = False
	closed = False
	requesteddate = DateNullCheck(Date())
	requestedtime = FixTime(TimeNullCheckAT(Time()))
	requestedinitials = ""
	issueddate = DateNullCheck(Date())
	issuedtime = FixTime(TimeNullCheckAT(Time()))
	issuedinitials = ""
	respondeddate = DateNullCheck(Date())
	respondedtime = FixTime(TimeNullCheckAT(Time()))
	respondedinitials = ""
	completeddate = DateNullCheck(Date())
	completedtime = FixTime(TimeNullCheckAT(Time()))
	completedinitials = ""
	finalizeddate = DateNullCheck(Date())
	finalizedtime = FixTime(TimeNullCheckAT(Time()))
	finalizedinitials = ""
	closeddate = DateNullCheck(Date())
	closedtime = FixTime(TimeNullCheckAT(Time()))
	closedinitials = ""
	laborreport = ""
	txtaccountpk = ""
	txtaccount = ""
	txtaccountdesc = ""
	txtcategorypk = ""
	txtcategory = ""
	txtcategorydesc = ""
	txtchargeable = False
	txtproblempk = ""
	txtproblem = ""
	txtproblemdesc = ""
	txtfailurepk = ""
	txtfailure = ""
	txtfailuredesc = ""
	txtsolutionpk = ""
	txtsolution = ""
	txtsolutiondesc = ""
	txtfailurewo = False
	txtmeter1reading = ""
	txtmeter2reading = ""
	txtisup = True
	assetexists = False
	ismeter = False
	txtAssetStatusHistoryPK = ""
	txtTaskInitials = ""
End If

txtFollowUpAllWO=False
txtFollowUpSingleWO=False
txtFollowUpMultiWO=False

If txtwogrouptype = "M" Then
	Mode = "WOGROUP"
End If

Select Case Mode
	Case "WO"
		DialogTitle = "Complete / Close Work Order"
		RightTitle = "Work Order #" & woid
	Case "WOGROUP"
		DialogTitle = "Complete / Close Work Orders"
		RightTitle = "Complete / Close Work Orders"
	Case Else
		DialogTitle = "Complete / Close Work Order"
		RightTitle = "Complete / Close Work Order"
End Select

Dim RS_WOClosePref
Dim WO_CLOSE_ACCOUNTSETALL
Dim WO_CLOSE_ALLTASKSCOMPLETE
Dim WO_CLOSE_CATEGORYSETALL
Dim WO_CLOSE_CHARGEABLE
Dim WO_CLOSE_CLOSE
Dim WO_CLOSE_COMPLETE
Dim WO_CLOSE_FINALIZE
Dim WO_CLOSE_LABORHOURSASN
Dim WO_CLOSE_LABORHOURSEST
Dim WO_CLOSE_LABORREPORT
Dim WO_CLOSE_MATERIALEST
Dim WO_CLOSE_OTHEREST
Dim WO_CLOSE_RESPOND
Dim WO_CLOSE_RETURNTOSERVICE
Dim WO_CLOSE_SETDOWNTIME
Dim WO_CLOSE_CUSTOMHOOK
Dim WO_CLOSE_FUSINGLE
Dim WO_CLOSE_FUMULTI

Set RS_WOClosePref = db.runSPReturnRS("MC_GetWorkOrderClosePrefs",Array(Array("@LaborPK", adInteger, adParamInput, 4, GetSession("USERPK")),Array("@RepairCenterPK", adInteger, adParamInput, 4, GetSession("RCPK"))),"")

Dim StatusDirectiveFromURL
StatusDirectiveFromURL = False

If UCase(Request.QueryString("responded")) = "Y" or _
   UCase(Request.QueryString("completed")) = "Y" or _
   UCase(Request.QueryString("finalized")) = "Y" or _
   UCase(Request.QueryString("closed")) = "Y" Then

	StatusDirectiveFromURL = True

End If

If db.dok Then
	If Not RS_WOClosePref.Eof Then

		If StatusDirectiveFromURL Then
			WO_CLOSE_RESPOND = False
			WO_CLOSE_COMPLETE = False
			WO_CLOSE_FINALIZE = False
			WO_CLOSE_CLOSE = False
		Else
			WO_CLOSE_RESPOND = RS_WOClosePref("WO_CLOSE_RESPOND")
			WO_CLOSE_FINALIZE = RS_WOClosePref("WO_CLOSE_FINALIZE")
			WO_CLOSE_COMPLETE = RS_WOClosePref("WO_CLOSE_COMPLETE")
			WO_CLOSE_CLOSE = RS_WOClosePref("WO_CLOSE_CLOSE")
		End If

		WO_CLOSE_ACCOUNTSETALL = RS_WOClosePref("WO_CLOSE_ACCOUNTSETALL")
		WO_CLOSE_ALLTASKSCOMPLETE = RS_WOClosePref("WO_CLOSE_ALLTASKSCOMPLETE")
		WO_CLOSE_CATEGORYSETALL = RS_WOClosePref("WO_CLOSE_CATEGORYSETALL")
		WO_CLOSE_CHARGEABLE = RS_WOClosePref("WO_CLOSE_CHARGEABLE")
		WO_CLOSE_LABORHOURSASN = RS_WOClosePref("WO_CLOSE_LABORHOURSASN")
		WO_CLOSE_LABORHOURSEST = RS_WOClosePref("WO_CLOSE_LABORHOURSEST")
		WO_CLOSE_LABORREPORT = NullCheck(RS_WOClosePref("WO_CLOSE_LABORREPORT"))
		WO_CLOSE_MATERIALEST = RS_WOClosePref("WO_CLOSE_MATERIALEST")
		WO_CLOSE_OTHEREST = RS_WOClosePref("WO_CLOSE_OTHEREST")
		WO_CLOSE_RETURNTOSERVICE = RS_WOClosePref("WO_CLOSE_RETURNTOSERVICE")
		WO_CLOSE_SETDOWNTIME = RS_WOClosePref("WO_CLOSE_SETDOWNTIME")
		WO_CLOSE_CUSTOMHOOK = NullCheck(RS_WOClosePref("WO_CLOSE_CUSTOMHOOK"))
		WO_CLOSE_FUSINGLE = NullCheck(RS_WOClosePref("WO_CLOSE_FUSINGLE"))
		WO_CLOSE_FUMULTI = NullCheck(RS_WOClosePref("WO_CLOSE_FUMULTI"))
	End If
End If
CloseObj RS_WOClosePref

If UCase(Request.QueryString("responded")) = "Y" Then
	WO_CLOSE_RESPOND = True
End If

If UCase(Request.QueryString("completed")) = "Y" Then
	WO_CLOSE_COMPLETE = True
End If

If UCase(Request.QueryString("finalized")) = "Y" Then
	WO_CLOSE_FINALIZE = True
End If

If UCase(Request.QueryString("closed")) = "Y" Then
	WO_CLOSE_CLOSE = True
End If
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">

<title><% =DialogTitle %></title>

<script language="javascript">
top.name = 'MCMENU';
</script>

<% If Application("SCRIPTENCODE") Then %>
<script LANGUAGE="JScript.Encode" SRC="../../javascript/encode/mc_util.jse"></script>
<script LANGUAGE="JScript.Encode" SRC="../../javascript/encode/mc_util_nodefer.jse"></script>
<% Else %>
<script LANGUAGE="JavaScript" SRC="../../javascript/normal/mc_util.js"></script>
<script LANGUAGE="JavaScript" SRC="../../javascript/normal/mc_util_nodefer.js"></script>
<script language="javascript" src="../../javascript/normal/mc_autocomplete.js"></script>
<% End If %>
<script src='../../javascript/jquery/jquery.1.4.3.js' type='text/javascript'></script>

<script type="text/javascript" language="javascript">
// RGJ V5 BEGIN
// Labor PK variable
var labor = '<%=GetSession("USERPK")%>';

// Repair Center PK variable
<%if WO_CLOSE_SHOWLABORACTUAL_FILTERRC = "Yes" Then%>
var repaircenter = '<%=GetSession("RCPK")%>';
<%Else%>
var repaircenter="";
<%End If%>

//alert(repaircenter);
//alert(labor);

$(document).ready(function(){
  
  $("#latype").change(function(){
    $("#labor").empty();
    getList("labor",null, $("#latype").val());
  });
  $("#partlocation").change(function(){
    $("#partContainer").hide();
    $("#partid").val("");
    getPartAutoComp($("#partid").val(), $("#partlocation").val());
  });

  $('#partid').bind({
    focus: function() {$(this).trigger('keyup');},
    keyup: function() {getPartAutoComp($(this).val(),$('#partlocation').val());}
  });

  $('body').click(function(e){
    var targetid = $(e.target).attr('id');
    targetid = targetid.toUpperCase();
    if (targetid=="PARTID"){
      $('#partContainer').show();
    } else {
      $('#partContainer').hide();
    }
  });
});


function getList(action, searchvalue, filtervalue) {
  //window.open("workorder_close_enhanced_data.asp?action=" + action + "&searchval=" + searchvalue + "&filterval=" + filtervalue + "&rc=" + repaircenter + "&la=" + labor);
  $.ajax({
    type: "POST",
    url: "workorder_close_enhanced_data.asp?action=" + action + "&searchval=" + searchvalue + "&filterval=" + filtervalue + "&rc=" + repaircenter + "&la=" + labor,
    dataType: "application/x-www-form-urlencoded",
    async: false,

    success: function(msg) {
      switch (action) {
        case "location":
          $("#partlocation").empty().append(msg);
          break;
        case "labor":
          $("#labor").empty().append(msg);
          break;
        //case "account":
        //  $("#account").empty().append(msg);
        //  break;
        //case "category":
        //  $("#category").empty().append(msg);
        //  break;
        //case "misccostaccount":
        //  $("#MiscCostAccount").empty().append(msg);
        //  break;
        //case "misccostcategory":
        //  $("#MiscCostCategory").empty().append(msg);
        //  break;
        //case "misccostvendor":
        //  $("#MiscCostVendor").empty().append(msg);
        //  break;
      }   // End Of Switch
    }   // End of Success
  })  // End of Ajax call
}   // End of Function getList()
function getPartAutoComp(searchvalue, filtervalue) {
  if ($('#partlocation').val()=="-1"){
    alert("You must select a location first.");
    return;
  }
  //window.open("workorder_close_enhanced_data.asp?action=inventory&searchval=" + searchvalue + "&filterval=" + filtervalue + "&rc=" + repaircenter);
  $.ajax({
    type: "POST",
    url: "workorder_close_enhanced_data.asp?action=inventory&searchval=" + searchvalue + "&filterval=" + filtervalue + "&rc=" + repaircenter,
    dataType: "html",
    async: false,

    success: function(msg) {
      try {
        $("#partContainer").html(msg);
        $("#partContainer").show();
      } catch(e){}
    }  // End of Success
  }); // End of Ajax call
}  // End of Function getPartAutoComp()
// RGJ V5 END


function processdefaults()
{
	var f = document.mcform;
	<% If WO_CLOSE_CLOSE and AccessToClose Then %>
	  processaction(document.images.closedimg,true)
	<% ElseIf WO_CLOSE_FINALIZE and AccessToFinalize Then %>
	processaction(document.images.finalizedimg,true)
	<% ElseIf WO_CLOSE_COMPLETE and AccessToComplete Then %>
	processaction(document.images.completedimg,true)
	<% ElseIf WO_CLOSE_RESPOND and AccessToRespond Then %>
	processaction(document.images.respondedimg,true)
	<% End If %>
	<% If Not WO_CLOSE_LABORREPORT = "" Then %>
	if (f.txtreport.value == ''){f.txtreport.value = '<% =JSEncode(WO_CLOSE_LABORREPORT) %>';}
	<% End If %>
	<% If WO_CLOSE_ACCOUNTSETALL Then %>
		f.txtAccountAll.checked = true;
	<% End If %>
	<% If WO_CLOSE_CHARGEABLE Then %>
		f.txtChargeable.checked = true;
	<% End If %>
	<% If WO_CLOSE_CATEGORYSETALL Then %>
		f.txtCategoryAll.checked = true;
	<% End If %>
	<% If WO_CLOSE_ALLTASKSCOMPLETE Then %>
		f.txtTasks.checked = true;
	<% End If %>

	<% If WO_CLOSE_FUSINGLE Then %>
		f.txtFollowUpSingleWO.checked = true;
		f.txtFollowUpMultiWO.checked = false;
		f.txtFollowUpAllWO.checked= true;
		f.txtFollowupChoice.value = '1';
	<% End If %>
	<% If WO_CLOSE_FUMULTI Then %>
		f.txtFollowUpSingleWO.checked = false;
		f.txtFollowUpMultiWO.checked = true;
		f.txtFollowUpAllWO.checked= true;
		f.txtFollowupChoice.value = '2';
	<% End If %>

	<% If WO_CLOSE_LABORHOURSEST Then %>
		f.txtLabor3.checked = true;
	<% End If %>
	<% If WO_CLOSE_LABORHOURSASN Then %>
		f.txtLabor1.checked = true;
	<% End If %>
	<% If WO_CLOSE_MATERIALEST Then %>
		f.txtMaterials.checked = true;
	<% End If %>
	<% If WO_CLOSE_OTHEREST Then %>
		f.txtOtherCost.checked = true;
	<% End If %>
	<% If WO_CLOSE_RETURNTOSERVICE Then %>
		if (top.trim(self.option_downtime2.innerText.toUpperCase()) != 'SET DOWNTIME')
		{
			f.txtDownTime.checked = true;
		}
	<% End If %>
	<% If WO_CLOSE_SETDOWNTIME Then %>
		if (top.trim(self.option_downtime2.innerText.toUpperCase()) == 'SET DOWNTIME')
		{
			f.txtDownTime.checked = true;
		}
	<% End If %>
}

var applyit = false;
var wopk = '<% =wopk %>';
var woid = '<% =woid %>';
var wogrouppk = '<% =wogrouppk %>';
var assetpk = '<% =assetpk %>';
var assetstatushistorypk = '<% =txtAssetStatusHistoryPK %>';
var mode = '<% =mode %>';
var myinitials = '<% =GetSession("UserInitials") %>';

function checkfollowup(o)
{

    //alert(o.value);
    var f = document.mcform;

    if (f.txtFollowUpAllWO.checked == true)
    {
        if (o.value == '1')
        {
            f.txtFollowUpSingleWO.checked = true;
            f.txtFollowUpMultiWO.checked = false;
        }
        else
        {
            f.txtFollowUpSingleWO.checked = false;
            f.txtFollowUpMultiWO.checked = true;
        }
    }
    else
    {
            f.txtFollowUpSingleWO.checked = false;
            f.txtFollowUpMultiWO.checked = false;
    }
}

function window_load()
{
  checkAssignments();
	<% If GetSession("smallres") = "Y" Then %>
	document.body.scroll='auto';
	<% End If %>

	var f = top.document.mcform;
	var fr = top;
	if (thetop.fraTopic && thetop.fraTopic.document)
	{
		var tf = thetop.fraTopic.document.mcform;
	}
	else
	{
		var tf = null;
	}
	var tfr = thetop.fraTopic;

	<% If (wogroupischecked or txtwogrouptype = "M") and ((Not WOGroupPK = "") and (Not WOGroupPK = "-1")) Then %>
	f.txtWOGroupAll.checked = true;
	wogrouptoggle();
	<% End If %>

    thetop.ismodal = true;

    processdefaults();
    enabledisable();

    try {
    if (thetop.acom == true) {
    var txtCategory_AC = new actb('CA','WOCLOSE',document.mcform.txtCategory);
    var txtAccount_AC = new actb('AC','WOCLOSE',document.mcform.txtAccount);
    var txtProblem_AC = new actb('FA_PROBLEM','WOCLOSE',document.mcform.txtProblem);
    var txtFailure_AC = new actb('FA_FAILURE','WOCLOSE',document.mcform.txtFailure);
    var txtSolution_AC = new actb('FA_SOLUTION','WOCLOSE',document.mcform.txtSolution);
    }
    } catch(e) {};
}

function window_unload()
{
	try
	{
		thetop.ismodal = false;
	}
	catch(e)
	{}
}

function afterpost(thecode)
{
	thetop.ismodal = false;

	if (thecode != null && thecode != '')
	{
		if (thetop.timeoutid != null)
		{
			thetop.clearTimeout(thetop.timeoutid);
		}

		thetop.timeoutid = thetop.setTimeout(thecode,10);
	}
}

var txtChargeable_old;
var txtFailedWO_old;
var option_meters_old;
var option_downtime_old;
var option_downtime2_old;
var downtimetable_old;
var mode_old;

function wogrouptoggle()
{
	f = top.document.mcform;

	if (f.txtWOGroupAll.checked == true)
	{
		top.wono.innerText = 'Work Order Group #' + wogrouppk;
		top.option_lri.style.display = '';

		txtChargeable_old = f.txtChargeable.checked;
		f.txtChargeable.checked = false;

		txtFailedWO_old = f.txtFailedWO.checked;
		f.txtFailedWO.checked = false;

		option_meters_old = top.option_meters.style.display;
		top.option_meters.style.display = 'none';

		option_downtime_old = top.option_downtime.style.display;
		top.option_downtime.style.display = '';

		option_downtime2_old = top.option_downtime2.innerText;
		top.option_downtime2.innerHTML = 'Return Asset to Service (if Shutdown)';

		downtimetable_old = top.downtimetable.style.width;
		top.downtimetable.style.width = '95%';

		tblSplitLabor.style.display='';
		requestedbox.style.display = 'none';
		issuedbox.style.display = 'none';
		WOAssignmentsBox.style.display = 'none';
		CompleteAssignmentsText.innerText = ' Set All Assignments Completed';
		CompleteAssignmentsBox.style.borderBottom = '0 solid #A2A2A2';
		mode_old = mode;
		f.mode.value = 'WOGROUP';
	}
	else
	{
		top.wono.innerText = 'Work Order #' + woid;
		top.option_lri.style.display = 'none';

		f.txtChargeable.checked = txtChargeable_old;
		f.txtFailedWO.checked = txtFailedWO_old;
		top.option_meters.style.display = option_meters_old;
		top.option_downtime.style.display = option_downtime_old;
		top.option_downtime2.innerText = option_downtime2_old;
		top.downtimetable.style.width = downtimetable_old;

    tblSplitLabor.style.display='none';
		requestedbox.style.display = '';
		issuedbox.style.display = '';
		WOAssignmentsBox.style.display = '';
		CompleteAssignmentsText.innerText = ' Completed Assignments:';
		CompleteAssignmentsBox.style.borderBottom = '1 solid #A2A2A2';

		f.mode.value = mode_old;
	}
}

function enabledisable(onload)
{
	if (onload == null)
	{
		onload = false;
	}

	f = top.document.mcform;

	if (f.txtAccount.value == '')
	{
		f.txtAccountAll.checked = false;
	}
	else
	{
		f.txtAccountAll.disabled = false;
	}

	if (f.txtCategory.value == '')
	{
		f.txtCategoryAll.checked = false;
	}
	else
	{
		f.txtCategoryAll.disabled = false;
	}

	if (f.txtTasks.checked == true)
	{
		f.txtTaskInitials.disabled = false;
		f.txtTaskInitials.className='required';
		if (onload == false)
		{
			try {
			//f.txtTaskInitials.focus();
			}
			catch(e) {}
		}
	}
	else
	{
		f.txtTaskInitials.disabled = true;
		f.txtTaskInitials.className='normal';
	}
}

function recorddowntime()
{
	if (assetpk == '')
	{
		return;
	}
	var param = new Object();
	param.caller = top;
	var mode = 'INSERVICE';
	var h = 545;
	param.mode = mode;
	param.downtimeonly = 'Y';

	var url = 'modules/common/mc_dialogrefreshable.asp?title=Asset+Downtime&url='+escape(thetop.path + thetop.oAS.rooturl + 'asset_statuschange.asp?assetpk='+ assetpk + '&mode=' + mode + '&wopk=' + wopk + '&downtimeonly=Y'+ '&assetstatushistorypk='+assetstatushistorypk);
	var param = top.loadmodal(url,param,430,h,false,false);
	if (param.cancel == true)
	{
		return false;
	}
	else
	{
		return true;
	}
}

function doapply(obj)
{
	f = top.document.mcform;

	if (obj.className.toLowerCase() == 'buttonsdisabled')
	{
		return;
	}

	obj.className = 'buttonsdisabled';
	document.images.cancelbutton.className = 'buttonsdisabled';

	if (validateform() == true)
	{
		if (mode == 'WO' && f.txtDownTime && f.txtDownTime.checked == true && assetpk != '')
		{
			if (recorddowntime() == true)
			{
				top.applyit = true;
				obj.className = 'buttonsdisabled';
				document.mcform.submit();
				return true;
			}
		}
		else
		{
			top.applyit = true;
			obj.className = 'buttonsdisabled';
			document.mcform.submit();
			return true;
		}
	}
	top.applyit = false;
	obj.className = 'buttonsenabled';
	document.images.cancelbutton.className = 'buttonsenabled';
	return false;
}

function validateform()
{
	if (top.standardvalidation_modal(document.mcform) == false)
	{
		top.mcinfo(null,errormessage);
		if (errorfield) {
			try {
				errorfield.focus();
			}
			catch(e)
			{
				// the field is not accessible or not visible so
				// must not be required
			}
		}
		return false;
	}
	return true;
}

function validateTextarea(ta){
	if (ta.tagName == "TEXTAREA"){
		if (ta.value.length > ta.maxlength) {
			alert("This field may only contain " + ta.maxlength + " characters and currently has " + ta.value.length + " characters.");
			ta.focus();
		}
	}
}

function checkAssignments(){
    var objelements = getElementsByClass("laborassign", null, "input");
    var isChecked = true;

    for(i=0; i<objelements.length; i++){
      var el = objelements[i];
      if (el.checked==false){
        isChecked = false;
      }
    }
    // if all addignments are checked then mark the Completed Assignemnts checkbox
    if (isChecked==true){
      document.mcform.CompleteAssignments.checked=true;
    }
}

function processit(obj){
	<%if WO_CLOSE_REQ_ASSIGNMENTS_FOR_CLOSE = "Yes" Then%>

    //Loop through any assignments and check to see if they are checked
    //getElementsByClass(searchClass,node,tag)
    var objelements = getElementsByClass("laborassign", null, "input");
    var isChecked = true;

    for(i=0; i<objelements.length; i++){
      var el = objelements[i];
      if (el.checked==false){
        isChecked = false;
      }
    }

    // if all addignments are checked then mark the Completed Assignemnts checkbox
    if (isChecked==true){
      document.mcform.CompleteAssignments.checked=true;
    }

	  //Return messages if an assignments checkbox is unchecked
	  if (isChecked==false){
      if (obj.name.toUpperCase()=="COMPLETEDIMG"){
        alert("All assignments must be complete before completing this work order.");
      }
      if (obj.name.toUpperCase()=="FINALIZEDIMG"){
        alert("All assignments must be complete before finalizing this work order.");
      }
      if (obj.name.toUpperCase()=="CLOSEDIMG"){
        alert("All assignments must be complete before closing this work order.");
      }
    } else {
      processaction(obj);
    }
  <%Else%>
    processaction(obj);
  <%End If%>

}

function processaction(obj,setdefault)
{
	if (setdefault == null)
	{
		setdefault = false;
	}

	switch (obj.name.toUpperCase())
	{
		case "RESPONDEDIMG":
		{
			if (document.mcform.txtResponded.value.toUpperCase() == 'Y')
			{
				if (setdefault == false)
				{
					a_close(true);
					a_finalize(true);
					a_complete(true);
					document.mcform.CompleteAssignments.checked = false;
					complete_assignments();
					a_respond(true);
				}
			}
			else
			{
				a_respond(false);
			}
			break;
		}
		case "COMPLETEDIMG":
		{
			if (document.mcform.txtCompleted.value.toUpperCase() == 'Y')
			{
				if (setdefault == false)
				{
					a_close(true);
					a_finalize(true);
					a_complete(true);
					document.mcform.CompleteAssignments.checked = false;
					complete_assignments();
				}
			}
			else
			{
				a_respond(false);
				a_complete(false);
				document.mcform.CompleteAssignments.checked = true;
				complete_assignments();
			}
			break;
		}
		case "FINALIZEDIMG":
		{
			if (document.mcform.txtFinalized.value.toUpperCase() == 'Y')
			{
				if (setdefault == false)
				{
					a_close(true);
					a_finalize(true);
				}
			}
			else
			{
				a_respond(false);
				a_complete(false);
				document.mcform.CompleteAssignments.checked = true;
				complete_assignments();
				a_finalize(false);
			}
			break;
		}
		case "CLOSEDIMG":
		{
			if (document.mcform.txtClosed.value.toUpperCase() == 'Y')
			{
				if (setdefault == false)
				{
					a_close(true);
				}
			}
			else
			{
				a_respond(false);
				a_complete(false);
				document.mcform.CompleteAssignments.checked = true;
				complete_assignments();
				a_finalize(false);
				a_close(false);
			}
			break;
		}

		default:
		{
			break;
		}
	}
}

function a_respond(b)
{
	if (b == true)
	{
		document.mcform.txtResponded.value = 'N';
		document.images.respondedimg.src = '../../images/button_respond_off.gif';
		respondeddata.style.display = 'none';
	}
	else
	{
		document.mcform.txtResponded.value = 'Y';
		document.images.respondedimg.src = '../../images/button_respond_on.gif';
		respondeddata.style.display = '';
		if (document.mcform.txtRespondedInitials.value == '')
		{
			document.mcform.txtRespondedInitials.value = myinitials;
		}
	}
}

function a_complete(b)
{
	if (b == true)
	{
		document.mcform.txtCompleted.value = 'N';
		document.images.completedimg.src = '../../images/button_complete_off.gif';
		completeddata.style.display = 'none';
	}
	else
	{
		document.mcform.txtCompleted.value = 'Y';
		document.images.completedimg.src = '../../images/button_complete_on.gif';
		completeddata.style.display = '';
		if (document.mcform.txtCompletedInitials.value == '')
		{
			document.mcform.txtCompletedInitials.value = myinitials;
		}
	}
}

function a_finalize(b)
{
	if (b == true)
	{
		document.mcform.txtFinalized.value = 'N';
		document.images.finalizedimg.src = '../../images/button_finalize_off.gif';
		finalizeddata.style.display = 'none';
	}
	else
	{
		document.mcform.txtFinalized.value = 'Y';
		document.images.finalizedimg.src = '../../images/button_finalize_on.gif';
		finalizeddata.style.display = '';
		if (document.mcform.txtFinalizedInitials.value == '')
		{
			document.mcform.txtFinalizedInitials.value = myinitials;
		}
	}
}

function a_close(b)
{
	if (b == true)
	{
		document.mcform.txtClosed.value = 'N';
		document.images.closedimg.src = '../../images/button_close_off.gif';
		closeddata.style.display = 'none';
	}
	else
	{
		document.mcform.txtClosed.value = 'Y';
		document.images.closedimg.src = '../../images/button_close_on.gif';
		closeddata.style.display = '';
		if (document.mcform.txtClosedInitials.value == '')
		{
			document.mcform.txtClosedInitials.value = myinitials;
		}
	}
}

function complete_assignments()
{
	var element;
	var f = document.mcform.elements;
	var b = document.mcform.CompleteAssignments.checked;

	// For each form element, extract the name and value
	for (var i = 0; i < f.length; i++) {
		element = f[i];

		if (element.type == "checkbox" && element.name.indexOf('CA_') != -1)
		{
			if (b == true)
			{
				element.checked = true;
			}
			else
			{
				try
				{
					element.checked = eval(element.oldcheckedvalue);
				}
				catch(e)
				{
					element.checked = false;
				}
			}
		}
	}
}

function openForm(url,w,h,l,t)
{
	<% Call SetSession("tm",Now) %>
	url = url + '&s=<% =BFEncrypt(SessionVars) %>';

	var aW = self.screen.availWidth-10;
	var aH = self.screen.availHeight-30;

	var aL = 0;
	var aT = 0;

	// for when using window.showmodeless or window.showmodal
	// aW = self.screen.availWidth;
	// aH = self.screen.availHeight;

	if (aW >= ((1024*2)-50) || aH >= ((768*2)-50))
	{
	 	// Using dual montior display
	 	if (aW >= ((1024*2)-50))
	 	{
	 		aW = aW / 2;
	 	}
	 	else
	 	{
	 		aH = aH / 2;
	 	}
	}

	if (w != null)
	{
		aW = w;
	}

	if (h != null)
	{
		aH = h;
	}

	if (l != null)
	{
		aL = l;
	}

	if (t != null)
	{
		aT = t;
	}

	try {
	self.closeformexternal = self.open(url,'closeformexternal',"width="+aW+",height="+aH+",left="+aL+",top="+aT+",scrollbars=auto,resizable=no,status=no,toolbar=no,menubar=no,location=no,directories=no");
	closeformexternal.focus();
	}
	catch(e) {};
}

function doc_keydown()
{
	if (self.event)
	{
		e = self.event;

		key = top.GetKeyString(e);

		if (key == 'Ctrl-Enter')
		{
			key = 'Alt-Enter';
		}

		//alert(key);
		switch(key)
		{
			case 'Alt-Enter':
			{
                top.applyit = true;
                top.doapply(document.images.applybutton);
				// cancel default action for key
				e.returnValue = false;
				// cancel bubble
				e.cancelBubble = true;
				return false;
				break;
			}
			default:
			{
				break;
			}
		}

	}

	return true;
}

//RGJ START
//material actual
function addMaterialRecord(){
  var fmsg = "The following fields require data entry:\n";
  var ffld = "";
  if (document.mcform.partlocation.value == -1){
    fmsg= fmsg + " - Stock Room Location \n";
    ffld = "partlocation";
//    alert("You must enter a Stock Room Location.");
 //   document.mcform.partlocation.focus();
//    return;
  }
  if (document.mcform.partid.value.length==0){
    fmsg = fmsg + " - Part \n";
    if (ffld.length==0){ffld = "partid";}
//    alert("You must enter a part.");
 //   document.mcform.partid.focus();
 //   return;
  }
  if (document.mcform.txtQuantity.value.length==0){
    fmsg = fmsg + " - Quantity \n";
    if (ffld.length==0){ffld = "txtQuantity";}
//    alert("You must enter a quantity of parts.");
//    document.mcform.txtQuantity.focus();
//    return;
  }
  if (document.mcform.deliverylocation.value.length==0){
    fmsg = fmsg + " - Delivery Location \n";
    if (ffld.length==0){ffld = "deliverylocation";}
  }
  if (ffld.length > 0){
    alert(fmsg);
    eval("document.mcform." +ffld+ ".focus()");
    event.cancelBubble = true;
    return;
  }

  var valuetovalidate;
  valuetovalidate = $("#partpk").val();
  //alert(valuetovalidate);
  if( valuetovalidate != "undefined" && valuetovalidate !== null && valuetovalidate != ""){
    //window.open("workorder_close_enhanced_data.asp?action=validate&searchval=" + valuetovalidate + "&filterval=IN");
    $.ajax({
      type: "POST",
      url: "workorder_close_enhanced_data.asp?action=validate&searchval=" + valuetovalidate + "&filterval=IN"  ,
      dataType: "html",
      async: false,

      success: function(msg) {
        try {
          if(msg=="-1"){
            alert("Invalid Part ID or Name. Please select an item from the list.");
            $("#partid").val("");
            $("#partid").focus();
            return;
          }
        } catch(e){}
      }  // End of Success
    }); // End of Ajax call
  } else {
    alert("Invalid Part ID or Name. Please select an item from the list.");
    $("#partid").val("");
    $("#partid").focus();
    return;    
  }


  var pk = "-1&"+ Date();
  var lpk = document.mcform.partlocation.value;
  var lnm = document.mcform.partlocation[document.mcform.partlocation.selectedIndex].text;

  //var typart = document.mcform.part.value;
  //var ppk = typart.substring( typart.indexOf("|")+1, typart.indexOf("~") );
  //var pid = typart.substring( typart.indexOf("~")+1, typart.length );
  var ppk = document.mcform.partpk.value;
  var pnm = document.mcform.partid.value;
  var qty = document.mcform.txtQuantity.value;
  var apk = document.mcform.account.value;
  if (apk==""){apk="-1";}
  var cpk = document.mcform.category.value;
  if (cpk==""){cpk="-1";}
  var dlval = document.mcform.deliverylocation.value.replace(/,/g, "");   // deliverylocation name
  
  var imgUnChk = '<img src="<%=Application( "web_path") & Application( "mapp_path" )%>images/checkbox_notchecked.jpg" border="0">&nbsp;';
  var imgChk = '<img src="<%=Application( "web_path") & Application( "mapp_path" )%>images/taskchecked.gif" border="0">';
//  var imgUnChk = '<input type="checkbox" disabled="disabled" >';
 // var imgChk = '<input type="checkbox" disabled="disabled" checked="checked">';
  var expedite = document.mcform.txtUdfBit2.checked ? 1 : 0;
  var expDisplay = expedite==1 ? imgChk : imgUnChk;
  var orditem = document.mcform.txtUdfBit1.checked ? 1 : 0;
  var ordDisplay = orditem==1 ?  imgChk : imgUnChk;
  var comms = document.mcform.txtPartComments.value.replace(/,/g, "");

  //var oTable  = top.findTable(oTBody);
  //var mydiv = oTable.id + 'popup';
  //var mytable = oTable.id + 'poptable';

  top.builddatarow(oma3body,2,null,pk,ppk,'WO',false,null,null,null,pnm,lnm,qty,'&nbsp;',expDisplay,ordDisplay);

  if (document.getElementById('oma3').style.display=='none'){
    document.getElementById("oma3").style.display='';
  }

  //Create an input type dynamically.
  var ACTPel = document.createElement("input");
  var LOCel = document.createElement("input");
  var PTel = document.createElement("input");
  var QTYel = document.createElement("input");
  
  var theKey = "x" + pk

  ACTPel.setAttribute("type", "hidden");
  ACTPel.setAttribute("value", "NEW");
  ACTPel.setAttribute("id", "PartAction");
  ACTPel.setAttribute("name", "PartAction");
  ACTPel.setAttribute("mcpk", theKey);

  LOCel.setAttribute("type", "hidden");
  LOCel.setAttribute("value", lpk);
  LOCel.setAttribute("id", "LocationPK");
  LOCel.setAttribute("name", "LocationPK");
  LOCel.setAttribute("mcpk", theKey);

  PTel.setAttribute("type", "hidden");
  PTel.setAttribute("value", ppk);
  PTel.setAttribute("id", "PartPK");
  PTel.setAttribute("name", "PartPK");
  PTel.setAttribute("mcpk", theKey);

  QTYel.setAttribute("type", "hidden");
  QTYel.setAttribute("value", qty);
  QTYel.setAttribute("id", "Quantity");
  QTYel.setAttribute("name", "Quantity");
  QTYel.setAttribute("mcpk", theKey);

  mcform.appendChild(ACTPel);
  mcform.appendChild(LOCel);
  mcform.appendChild(PTel);
  mcform.appendChild(QTYel);

  //if (apk!==""){
   var ACCTel = document.createElement("input");
   ACCTel.setAttribute("type", "hidden");
   ACCTel.setAttribute("value", apk);
   ACCTel.setAttribute("id", "AccountID");
   ACCTel.setAttribute("name", "AccountID");
   ACCTel.setAttribute("mcpk", theKey);
   mcform.appendChild(ACCTel);
  //}
  //if (cpk!==""){
   var CATel = document.createElement("input");
   CATel.setAttribute("type", "hidden");
   CATel.setAttribute("value", cpk);
   CATel.setAttribute("id", "CategoryID");
   CATel.setAttribute("name", "CategoryID");
   CATel.setAttribute("mcpk", theKey);
   mcform.appendChild(CATel);
  //}
  
   var DLel = document.createElement("input");
   DLel.setAttribute("type", "hidden");
   DLel.setAttribute("value", dlval);
   DLel.setAttribute("id", "UDFChar1");
   DLel.setAttribute("name", "UDFChar1");
   DLel.setAttribute("mcpk", theKey);
   mcform.appendChild(DLel);
  
   var EXPel = document.createElement("input");
   EXPel.setAttribute("type", "hidden");
   EXPel.setAttribute("value", expedite);
   EXPel.setAttribute("id", "UDFBit2");
   EXPel.setAttribute("name", "UDFBit2");
   EXPel.setAttribute("mcpk", theKey);
   mcform.appendChild(EXPel);

   var ORDel  = document.createElement("input");
   ORDel.setAttribute("type", "hidden");
   ORDel.setAttribute("value", orditem);
   ORDel.setAttribute("id", "UDFBit1");
   ORDel.setAttribute("name", "UDFBit1");
   ORDel.setAttribute("mcpk", theKey);
   mcform.appendChild(ORDel);
   
   var COMel = document.createElement("input");
   COMel.setAttribute("type", "hidden");
   COMel.setAttribute("value", comms);
   COMel.setAttribute("id", "WOP_Comments");
   COMel.setAttribute("name", "WOP_Comments");
   COMel.setAttribute("mcpk", theKey);
   mcform.appendChild(COMel);
   
   hidePopUp('INV');
}

//labor actuals
function addLaborRecord(){
  if (document.mcform.labor.value.length==0){
    alert("You must enter a labor.");
    document.mcform.labor.focus();
    return;
  }
  if (document.mcform.txtWorkDate.value.length==0){
    alert("You must enter a work date.");
    document.mcform.txtWorkDate.focus();
    return;
  }
  if (document.mcform.txtHoursRegular.value.length==0 && document.mcform.txtHoursOvertime.value.length==0 && document.mcform.txtHoursOther.value.length==0){
    alert("You must enter Regular, Overtime, or Other hours.");
    document.mcform.txtHoursRegular.focus();
    return;
  }

  var pk = "-1&"+ Date();
  var typlab = document.mcform.labor.value;
  var lpk = typlab.substring( typlab.indexOf("|")+1, typlab.length );
  var wdate = document.mcform.txtWorkDate.value;
  var lnm = document.mcform.labor[document.mcform.labor.selectedIndex].text;
  var rh = document.mcform.txtHoursRegular.value;
  var oh = document.mcform.txtHoursOvertime.value;
  var xh = document.mcform.txtHoursOther.value;
  var lcom = document.mcform.txtLaborComments.value.replace(/,/g, "");
  //lcom = lcom.replace(/,/g,"");

  if (rh==""){
    rh="0";
  }else{
    rh=Math.round(rh*100)/100;
    rh = rh+'';
  }
  if (oh==""){
    oh="0";
  }else{
    oh=Math.round(oh*100)/100;
    oh = oh+''
  }
  if (xh==""){
    xh="0";
  }else{
    xh=Math.round(xh*100)/100;
    xh = xh+''
  }


  top.builddatarow(ola3body,2,null,pk,lpk,'WO',false,null,null,null,wdate,lnm,rh,oh,xh);

  if (document.getElementById('ola3').style.display=='none'){
    document.getElementById("ola3").style.display='';
  }

  //Create an input type dynamically.
  var ACTLel = document.createElement("input");
  var LAel = document.createElement("input");
  var RTel = document.createElement("input");
  var OTel = document.createElement("input");
  var XTel = document.createElement("input");
  var WDel = document.createElement("input");
  var LCel = document.createElement("input");
  var theKey = "x" + pk

  ACTLel.setAttribute("type", "hidden");
  ACTLel.setAttribute("value", "NEW");
  ACTLel.setAttribute("id", "LaborAction");
  ACTLel.setAttribute("name", "LaborAction");
  ACTLel.setAttribute("mcpk", theKey);

  LAel.setAttribute("type", "hidden");
  LAel.setAttribute("value", lpk);
  LAel.setAttribute("id", "LaborPK");
  LAel.setAttribute("name", "LaborPK");
  LAel.setAttribute("mcpk", theKey);

  RTel.setAttribute("type", "hidden");
  RTel.setAttribute("value", rh);
  RTel.setAttribute("id", "RegularHours");
  RTel.setAttribute("name", "RegularHours");
  RTel.setAttribute("mcpk", theKey);

  OTel.setAttribute("type", "hidden");
  OTel.setAttribute("value", oh);
  OTel.setAttribute("id", "OvertimeHours");
  OTel.setAttribute("name", "OvertimeHours");
  OTel.setAttribute("mcpk", theKey);

  XTel.setAttribute("type", "hidden");
  XTel.setAttribute("value", xh);
  XTel.setAttribute("id", "Otherhours");
  XTel.setAttribute("name", "Otherhours");
  XTel.setAttribute("mcpk", theKey);

  WDel.setAttribute("type", "hidden");
  WDel.setAttribute("value", wdate);
  WDel.setAttribute("id", "WorkDate");
  WDel.setAttribute("name", "WorkDate");
  WDel.setAttribute("mcpk", theKey);

  LCel.setAttribute("type", "hidden");
  LCel.setAttribute("value", lcom);
  LCel.setAttribute("id","LaborComments");
  LCel.setAttribute("name","LaborComments");
  LCel.setAttribute("mcpk", theKey);;

  mcform.appendChild(ACTLel);
  mcform.appendChild(LAel);
  mcform.appendChild(RTel);
  mcform.appendChild(OTel);
  mcform.appendChild(XTel);
  mcform.appendChild(WDel);
  mcform.appendChild(LCel);

  hidePopUp('EMP');
}

//other misc cost actuals
function addMiscCostRecord(){
  if (document.mcform.txtMiscCostName.value.length==0){
    alert("You must enter a Miscellaneous Cost Name.");
    document.mcform.txtMiscCostName.focus();
    return;
  }
  if (document.mcform.txtMiscCostDate.value.length==0){
    alert("You must enter a Miscellaneous Cost Date.");
    document.mcform.txtMiscCostDate.focus();
    return;
  }

  var pk = "-1&"+ Date();
  var mcn = document.mcform.txtMiscCostName.value.replace(/,/g, "");
  mcn = mcn.replace(","," -")
  var mcd = document.mcform.txtMiscCostDesc.value.replace(/,/g, "");
  var mcdate = document.mcform.txtMiscCostDate.value;
  var mcv = document.mcform.MiscCostVendor.value;
  if (mcv==""){mcv="-1";}
  var mca = document.mcform.MiscCostAccount.value;
  if (mca==""){mca="-1";}
  var mcc = document.mcform.MiscCostCategory.value;
  if (mcc==""){mcc="-1";}
  var mcp = document.mcform.txtPrice.value;
  var mcom = document.mcform.txtMiscCostComments.value.replace(/,/g, "");
  //mcom = mcom.replace(/,/g,"");
  if (mcp==""){mcp="0.00";}

  top.builddatarow(oot3body,2,null,pk,null,'WO',false,null,null,null,mcdate,mcn,mcp);

  if (document.getElementById('oot3').style.display=='none'){
    document.getElementById("oot3").style.display='';
  }

  //Create an input type dynamically.
  var ACTMCel = document.createElement("input");
  var MCName = document.createElement("input");
  var MCDesc = document.createElement("input");
  var MCDate = document.createElement("input");
  var MCPrice = document.createElement("input");
  var MCComments = document.createElement("input");
  var theKey = "x" + pk

  ACTMCel.setAttribute("type", "hidden");
  ACTMCel.setAttribute("value", "NEW");
  ACTMCel.setAttribute("id", "MiscCostAction");
  ACTMCel.setAttribute("name", "MiscCostAction");
  ACTMCel.setAttribute("mcpk", theKey);

  MCName.setAttribute("type", "hidden");
  MCName.setAttribute("value", mcn);
  MCName.setAttribute("id", "txtMCName");
  MCName.setAttribute("name", "txtMCName");
  MCName.setAttribute("mcpk", theKey);

  MCDesc.setAttribute("type", "hidden");
  MCDesc.setAttribute("value", mcd);
  MCDesc.setAttribute("id", "txtMCDesc");
  MCDesc.setAttribute("name", "txtMCDesc");
  MCDesc.setAttribute("mcpk", theKey);

  MCDate.setAttribute("type", "hidden");
  MCDate.setAttribute("value", mcdate);
  MCDate.setAttribute("id", "txtMCDate");
  MCDate.setAttribute("name", "txtMCDate");
  MCDate.setAttribute("mcpk", theKey);

  MCPrice.setAttribute("type", "hidden");
  MCPrice.setAttribute("value", mcp);
  MCPrice.setAttribute("id", "txtMCPrice");
  MCPrice.setAttribute("name", "txtMCPrice");
  MCPrice.setAttribute("mcpk", theKey);

  MCComments.setAttribute("type", "hidden");
  MCComments.setAttribute("value", mcom);
  MCComments.setAttribute("id", "MiscCostComments");
  MCComments.setAttribute("name", "MiscCostComments");
  MCComments.setAttribute("mcpk", theKey);

  mcform.appendChild(ACTMCel);
  mcform.appendChild(MCName);
  mcform.appendChild(MCDesc);
  mcform.appendChild(MCDate);
  mcform.appendChild(MCPrice);
  mcform.appendChild(MCComments);

  //if (mcv!==""){
    var MCVendor = document.createElement("input");
    MCVendor.setAttribute("type", "hidden");
    MCVendor.setAttribute("value", mcv);
    MCVendor.setAttribute("id", "txtMCVendorID");
    MCVendor.setAttribute("name", "txtMCVendorID");
    MCVendor.setAttribute("mcpk", theKey);
    mcform.appendChild(MCVendor);
  //}
  //if (mca!==""){
    var MCAccount = document.createElement("input");
    MCAccount.setAttribute("type", "hidden");
    MCAccount.setAttribute("value", mca);
    MCAccount.setAttribute("id", "txtMCAccountID");
    MCAccount.setAttribute("name", "txtMCAccountID");
    MCAccount.setAttribute("mcpk", theKey);
    mcform.appendChild(MCAccount);
  //}
  //if (mcc!==""){
    var MCCategory = document.createElement("input");
    MCCategory.setAttribute("type", "hidden");
    MCCategory.setAttribute("value", mcc);
    MCCategory.setAttribute("id", "txtMCCategoryID");
    MCCategory.setAttribute("name", "txtMCCategoryID");
    MCCategory.setAttribute("mcpk", theKey);
    mcform.appendChild(MCCategory);
  //}


  hidePopUp('MISC');

}


//Deletes user added row
function mydeleterow(otable) {

	try {
	  if (top.CheckAccess('EDIT',thetop.currentmodule.toUpperCase() + 'E') == false){return true;}
	}
	catch(e) {}

  var numsel = 0;

  var rows=otable.rows;

  var oTBody = top.findTBody(rows[0]);

  var cannotdelete = false;

  for(var i=0;i<rows.length;i++)
  {
    if (rows[i].mcguid != null)
    {
		var cells=rows[i].cells;
		var kids = cells[0].children;

		  if(kids.length != 0){
			  if(kids[0].checked) {
	   			numsel++;

          //RGJ BEGIN
          var j
					var thepk = "x"+rows[i].mcguid;
					var txt="";
          var objelementsNamed = getElementsByAttribute(document.getElementById("mcform"), "input", "mcpk", thepk);

          for(j=0; j<objelementsNamed.length; j++){
            var el = objelementsNamed[j];
            mcform.removeChild(el);
          }
          //RGJ END

					rows[i].removeNode(true);
					i--
			  }
		  }
	  }
  }
  if (numsel == 0){
  	eval('thetop.table_' + otable.id.substring(1) + '_state = false');
	  top.mcinfo(null,"There were not any items checked. Please try again.");
	  return;
  }
}
//Function added for mydeleterow function above
function getElementsByAttribute(oElm, strTagName, strAttributeName, strAttributeValue){
	var arrElements = (strTagName == "*" && oElm.all)? oElm.all : oElm.getElementsByTagName(strTagName);
	var arrReturnElements = new Array();
	var oAttributeValue = (typeof strAttributeValue != "undefined")? new RegExp("(^|\\s)" + strAttributeValue + "(\\s|$)") : null;
	var oCurrent;
	var oAttribute;
	for(var i=0; i<arrElements.length; i++){
		oCurrent = arrElements[i];
		oAttribute = oCurrent.getAttribute && oCurrent.getAttribute(strAttributeName);
		if(typeof oAttribute == "string" && oAttribute.length > 0){
			if(typeof strAttributeValue == "undefined" || (oAttributeValue && oAttributeValue.test(oAttribute))){
				arrReturnElements.push(oCurrent);
			}
		}
	}
	return arrReturnElements;
}

//Function Added to find labor assigments.
function getElementsByClass(searchClass,node,tag) {
	var classElements = new Array();
	if ( node == null )
		node = document;
	if ( tag == null )
		tag = '*';
	var els = node.getElementsByTagName(tag);
	var elsLen = els.length;
	var pattern = new RegExp("(^|\\s)"+searchClass+"(\\s|$)");
	for (i = 0, j = 0; i < elsLen; i++) {
		if ( pattern.test(els[i].className) ) {
			classElements[j] = els[i];
			j++;
		}
	}
	return classElements;
}
//Function that fills in Current date in form date fields on double click
function GetCurrentDate(obj){
  var tz = "<%=TZ%>";
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1;//January is 0!
  var yyyy = today.getFullYear();
  //if(dd<10){dd='0'+dd}
  //if(mm<10){mm='0'+mm}
  if (tz=="Y"){
    var theDate = dd+'/'+mm+'/'+yyyy
  } else {
    var theDate = mm+'/'+dd+'/'+yyyy
  }
  obj.value=theDate;
}
//RGJ END

<% If errormessage = "" Then %>
window.onload = window_load;

<% End If %>
window.onunload = window_unload;
document.onclick = top.doc_click;
document.onkeydown = self.doc_keydown;
</script>

<script language="javascript" type="text/javascript">
//RGJ BEGIN - Cost Actual popup stuff
  var ie = document.all
  var ns6 = document.getElementById && !document.all

  function showPopUp(type) {
    var tz = "<%=TZ%>";
    var cd = new Date();
    var theDay = cd.getDate();
    var theMonth = cd.getMonth();
    theMonth++;
    var theYear = cd.getFullYear();

    if (tz == "Y") {
      var now = theDay + "/" + theMonth + "/" + theYear;
    } else {
      var now = theMonth + "/" + theDay + "/" + theYear;
    }
    if (type == "EMP") {
      document.getElementById('pWinEmp').style.visibility = 'visible';

      $("#latype").val("EMP");
      getList("labor", null, $("#latype").val());

      document.mcform.latype.focus();
      document.mcform.txtWorkDate.value = now;
    } else if (type == "INV") {
      document.mcform.partpk.value = "";
      document.mcform.partid.value = "";
      document.mcform.txtQuantity.value = "";
      document.mcform.deliverylocation.value = "";
      document.mcform.txtUdfBit1.checked = "";
      document.mcform.txtUdfBit2.checked = "";
      document.mcform.txtPartComments.value = "";
      
      document.getElementById('pWinInv').style.visibility = 'visible';

      getList("location");
      //$("#partlocation option")[2].attr('selected', 'selected')
      $("#partlocation option:nth-child(1)").attr("selected", "selected")
      //getList("inventory", $("#partlocation").val(), rcpk);
      //getList("account");
      //getList("category");

      document.mcform.partlocation.focus();
    } else if (type == "MISC") {
      document.getElementById('pWinMisc').style.visibility = 'visible';

      //getList("misccostaccount");
      //getList("misccostcategory");
      //getList("misccostvendor");
      document.mcform.txtMiscCostName.value = "";
      document.mcform.txtMiscCostDesc.value = "";
      document.mcform.MiscCostVendor.value = "";
      document.getElementById("MiscCostVendorDesc").innerHTML = "";
      document.mcform.txtMiscCostDate.value = "";
      document.mcform.MiscCostAccount.value = "";
      document.mcform.MiscCostCategory.value = "";
      document.mcform.txtPrice.value = "";
      document.mcform.txtMiscCostComments.value = "";


      document.mcform.txtMiscCostName.focus();
      document.mcform.txtMiscCostDate.value = now;
    }

  }
  function hidePopUp(type) {
    if (type=="EMP"){
      document.getElementById('pWinEmp').style.visibility = 'hidden';
    } else if (type=="INV"){
      document.getElementById('pWinInv').style.visibility = 'hidden';
    } else if (type == "MISC") {
      document.getElementById('pWinMisc').style.visibility = 'hidden';
    }
  }
  function startPopUp(popuptype) {
    if (popuptype == "INV") {
      timerID = setTimeout("showPopUp('INV')", 0);
    } else if (popuptype == "EMP") {
      timerID = setTimeout("showPopUp('EMP')", 0);
    } else if (popuptype == "MISC") {
      timerID = setTimeout("showPopUp('MISC')", 0);
    }
    //fillItems(popuptype);
  }
  /*
  heres where we set the popup window class properties..

  */
//RGJ END
</script>
<!-- //=============================// -->
<script language="javascript" type="text/javascript">
//Meter Reading Validation New reading cannot be less than current readings
function validateMeterReading(obj){
<%If WO_CLOSE_UPDATEASSETMETERS = "Yes" Then %>
  var objname = obj.name;
  var meterreading = obj.value;
  var oldmeter1reading = document.mcform.assetmeter1reading.value;
  var oldmeter2reading = document.mcform.assetmeter2reading.value;

  if (oldmeter1reading==null||oldmeter1reading==""){oldmeter1reading=0;}
  if (oldmeter2reading==null||oldmeter2reading==""){oldmeter2reading=0;}

  if (objname=="txtMeter1Reading"){
    if (parseInt(meterreading)<parseInt(oldmeter1reading)){
      top.mcinfo(null,"Please enter a value greater than the current meter 1 reading.");
      //var answer = confirm("The value you entered is less than the current meter 1 reading.  Continue?");
      //if (!answer){
        document.mcform.txtMeter1Reading.value=oldmeter1reading;
        document.mcform.txtMeter1Reading.focus();
        return false;
      //}
    }
  } else if (objname=="txtMeter2Reading"){
    if (parseInt(meterreading)<parseInt(oldmeter2reading)){
      top.mcinfo(null,"Please enter a value greater than the current meter 2 reading..");
      document.mcform.txtMeter2Reading.value=oldmeter2reading;
      document.mcform.txtMeter2Reading.focus();
      return false;
    }
  }
  document.mcform.txtMeter1Reading.focus();
<%End If %>
}
//Function checks to see how many Meter PM work order will be created with user entered meter reading
function checkMeterPMs(){

  // Meter readings from form
  var m1 = document.mcform.txtMeter1Reading.value;
  var m2 = document.mcform.txtMeter2Reading.value;

  // Values from Meter PMS
  var pm = "<%=MeterPMs%>";
  var pmi = "<%=Meter1PMIntervals%>";
  var pmn = "<%=Meter1PMNext%>";
  var pmi2 = "<%=Meter2PMIntervals%>";
  var pmn2 = "<%=Meter2PMNext%>";

  //other variables
  var cnt1=0;
  var cnt2=0;
  var cnt1Total=0;
  var cnt2Total=0;
  var avg1
  var avg2
  if (pmn==""){pnm=0}
  if (pmn2==""){pmn2=0}

  //Count WOs to be generated
  if (pm.indexOf(",")>0){ //multiple values
    //Create Arrays from Meter PM Data
    var pmArray = pm.split(",");
    var pmiArray = pmi.split(",");
    var pmnArray = pmn.split(",");
    var pmiArray2 = pmi2.split(",");
    var pmnArray2 = pmn2.split(",");

    //Loop through PMs
    for (var i=0;i<pmArray.length;i++){
      if (pmn[i]!=0){
        // subtract next PM interval (pmn(array[i]) from entered meter reading and divide by PM Interval
        cnt1 = (parseInt(m1) - parseInt(pmnArray[i])) / parseInt(pmiArray[i])+1 // (50-30)/10  & (50-50)/50
      }
      if (pmn2[i]!=0){
        cnt2 = (parseInt(m2) - parseInt(pmnArray2[i])) / parseInt(pmiArray2[i])+1
      }
      cnt1Total = parseInt(cnt1Total) + parseInt(cnt1);
      cnt2Total = parseInt(cnt2Total) + parseInt(cnt2);
    }
    // Math.round(num*Math.pow(10,dec))/Math.pow(10,dec)
    avg1 = Math.round((cnt1Total / pmArray.length)*Math.pow(10,2))/Math.pow(10,2);
    avg2 = Math.round((cnt2Total / pmArray.length)*Math.pow(10,2))/Math.pow(10,2);
  } else { //single or no value
    //Make sure there is a PM
    if (pm.length!=0){
      //Check for next interval > 0
      if (pmn!=0){
        cnt1 = ((parseInt(m1)-parseInt(pmn)) / parseInt(pmi))+1  // (30-10 / 10)+1
      }
      if (pmn2!=0){
        cnt2 = ((parseInt(m2)-parseInt(pmn2)) / parseInt(pmi2))+1
      }
      avg1=cnt1;
      avg2=cnt2;
    }
  }

  //If more than 1 PM WO then alert user
  if (avg1>1||avg2>1){
    var msg;
    msg = "The average number of Meter 1 based PM work orders created by this change is " + avg1.toString() + ".\n";
    if (avg2>0){
      msg = msg + "The average number of Meter 2 based PM work orders created by this change is " + avg2.toString() + ".\n";
    }
    msg = msg + "\nDo you want to continue?";

    var returnVal = confirm(msg);
    if (returnVal==true){// user is ok - process past to submit routine
      top.applyit = true; top.doapply(document.images.applybutton);
    } else {// user canceled out of message - stop and focus on Meter 1 reading field
      document.mcform.txtMeter1Reading.focus();
    }
  } else { // otherwise just process past and call the submit routine
    top.applyit = true; top.doapply(document.images.applybutton);
  }
}

//Called by Apply Button - Checks for required fields then calls the function above to check Meter PM work orders

function doChecks(){
  <%if TaskCode = -1 And WO_CLOSE_ALLTASKSCOMPLETE_REQ = "Yes" Then%>
    //see if user checked "Set All Tasks Complete"
    var atc = document.mcform.txtTasks.checked;
    
    if (document.mcform.txtCompleted.value.toUpperCase() == 'Y'){
      if(atc==false){ // "Set All Tasks Complete" was not checked and user is completing and all tasks are not complete
        top.mcinfo(null,'All Tasks must be complete before completing the work order.');
        return;
      }
    }
  <%End If%>

  <%If WO_CLOSE_SHOWLABORREPORT_REQ = "Yes" Then%>
    if (document.mcform.txtreport.value.length==0){
      top.mcinfo(null,'Labor Report Information is required.');
	    document.mcform.txtreport.focus();
      return;
    }
  <%End If
  If WO_CLOSE_SHOWACCOUNTCATEGORY_AREQ = "Yes" Then%>
    if (document.mcform.txtAccount.value.length==0){
      top.mcinfo(null,'Account is required.');
	    document.mcform.txtAccount.focus();
      return;
    }
  <%End If
  If WO_CLOSE_SHOWACCOUNTCATEGORY_CREQ = "Yes" Then%>
    if (document.mcform.txtCategory==0){
      top.mcinfo(null,'Category is required.');
	    document.mcform.txtCategory.focus();
      return;
    }
  <%End If
  If WO_CLOSE_SHOWMETERREADINGS_REQM1 = "Yes" Then%>
    if (document.mcform.txtMeter1Reading.value.length==0){
      top.mcinfo(null,'Meter 1 Reading is required.');
	    document.mcform.txtMeter1Reading.focus();
      return;
    }
  <%End If
  If WO_CLOSE_SHOWMETERREADINGS_REQM2 = "Yes" Then%>
    if (document.mcfrom.txtMeter2Reading.value.length==0){
      top.mcinfo(null,'Meter 2 Reading is required.');
	    document.mcform.txtMeter2Reading.focus();
      return;
    }
  <%End If
  If WO_CLOSE_SHOWFAILUREANALYSIS_PREQ = "A" Then%>
    if (document.mcform.txtProblem.value.length==0){
      top.mcinfo(null,'A Problem Code is required.');
	    document.mcform.txtProblem.focus();
	    return;
    }
  <%ElseIf WO_CLOSE_SHOWFAILUREANALYSIS_PREQ = "C" Then %>
    if (document.mcform.txtCompleted.value.toUpperCase() == 'Y'){
      if (document.mcform.txtProblem.value.length==0){
        top.mcinfo(null,'A Problem Code is required before completing the work order.');
	      document.mcform.txtProblem.focus();
	      return;
      }
    }
  <%End If 
  If WO_CLOSE_SHOWFAILUREANALYSIS_FREQ = "A" Then%>
    if (document.mcform.txtFailure.value.length==0){
      top.mcinfo(null,'A Failure Reason Code is required.');
	    document.mcform.txtFailure.focus();
      return;
    }
  <%ElseIf WO_CLOSE_SHOWFAILUREANALYSIS_FREQ = "C" Then %>
    if (document.mcform.txtCompleted.value.toUpperCase() == 'Y'){
      if (document.mcform.txtFailure.value.length==0){
        top.mcinfo(null,'A Failure Reason Code is required before completing the work order.');
	      document.mcform.txtFailure.focus();
        return;
      }
    }
  <%End If
  If WO_CLOSE_SHOWFAILUREANALYSIS_SREQ = "A" Then%>
    if (document.mcform.txtSolution.value.length==0){
      top.mcinfo(null,'A Solution Code is required.');
	    document.mcform.txtSolution.focus();
      return;
    }
  <%ElseIf WO_CLOSE_SHOWFAILUREANALYSIS_SREQ = "C" Then%>
    if (document.mcform.txtCompleted.value.toUpperCase() == 'Y'){
      if (document.mcform.txtSolution.value.length==0){
        top.mcinfo(null,'A Solution Code is required before completing the work order.');
	      document.mcform.txtSolution.focus();
        return;
      }
    }
  <%End If
  If Mode = "WO" AND WO_CLOSE_SHOWLABORACTUAL_REQ = "Yes" Then%>
    var ltbl = document.getElementById('ola3')
    var rcnt = ltbl.rows.length;
    if (rcnt < 5){
      top.mcinfo(null,'At least one Labor Actual is required.');
      return;
    }
  <%End If
  If Mode = "WO" AND WO_CLOSE_SHOWPARTACTUAL_REQ = "Yes" Then%>
    var mtbl = document.getElementById('oma3')
    var rcnt = mtbl.rows.length;
    if (rcnt < 5){
      top.mcinfo(null,'At least one Part Actual is required.');
      return;
    }
  <%End If
  If Mode = "WO" AND WO_CLOSE_SHOWMISCCOSTACTUAL_REQ = "Yes" Then%>
    var otbl = document.getElementById('oot3')
    var rcnt = otbl.rows.length;
    if (rcnt < 5){
      top.mcinfo(null,'At least one Miscellaneous Cost Actual is required.');
      return;
    }
  <%End If%>

  //Call validateMeterReading function
  try {
    if (validateMeterReading(document.mcform.txtMeter1Reading)==false){
      //document.mcform.txtMeter1Reading.focus();
      return;
    }
  } catch(e) {};
  try {
    if (validateMeterReading(document.mcform.txtMeter2Reading)==false){
      //document.mcform.txtMeter2Reading.focus();
      return;
    }
  } catch(e) {};

  // Call checkMeterPMs function
<%If WO_CLOSE_CHECKMETERREADINGDELTA = "Yes" Then%>
  //Check PMs and submit from that function
  checkMeterPMs();
<%Else%>
  //Submit form
  top.applyit = true; top.doapply(document.images.applybutton);
<%End If%>
}
function select_all(thetable){
	var de = document.mcform.elements;
	for (var i=0; i<de.length; i++){
	  if(de[i].name.indexOf("checkedDataList_" + thetable) != -1){
			if(de[i].name.indexOf("ROW_ID") == -1){
				try{
				  if (top.findRow(de[i]).style.display == ''){
					  de[i].checked = eval('!thetop.table_' + thetable + '_state');
					  top.mctr_checked(de[i]);
				  }
				} catch(e) {}
			}
		}
	}
	eval('thetop.table_' + thetable + '_state = !thetop.table_' + thetable + '_state');
}
function editCost(obj,type){
  var mcid = obj.mcguid;
  //alert(mcid);
  var rowid = obj.rowIndex;
  //alert(rowid);
  var partloc=obj.cells[2].innerHTML;
  var partid=obj.cells[1].innerHTML;
  var partqty=obj.cells[3].innerHTML;

  document.getElementById('partEditWindow').style.visibility = 'visible';
  document.getElementById('editPartLocation').innerHTML=partloc;
  document.getElementById('editPartID').innerHTML=partid;
  document.getElementById('editPartQuantity').value=partqty;
  document.getElementById('partRecordID').value=mcid;
  document.getElementById('partRowID').value=rowid;
  document.mcform.editPartQuantity.focus();
  document.mcform.editPartQuantity.select();
}
function updatePartRecord(trid,newqty,rowindex){
  // Set value on screen
  var x=document.getElementById('oma3').rows
  var y=x[rowindex].cells
  y[3].innerHTML=newqty;
  y[3].style.backgroundColor="yellow";

  //Create an input type dynamically.
  var ACTel = document.createElement("input");
  var RECel = document.createElement("input");
  var QTYel = document.createElement("input");

  ACTel.setAttribute("type", "hidden");
  ACTel.setAttribute("value", "UPD");
  ACTel.setAttribute("id", "PartAction2");
  ACTel.setAttribute("name", "PartAction2");
  ACTel.setAttribute("mcpk", trid);

  RECel.setAttribute("type", "hidden");
  RECel.setAttribute("value", trid);
  RECel.setAttribute("id", "WOPartPK");
  RECel.setAttribute("name", "WOPartPK");
  RECel.setAttribute("mcpk", trid);

  QTYel.setAttribute("type", "hidden");
  QTYel.setAttribute("value", newqty);
  QTYel.setAttribute("id", "UPDPartQTY");
  QTYel.setAttribute("name", "UPDPartQTY");
  QTYel.setAttribute("mcpk", trid);

  mcform.appendChild(ACTel);
  mcform.appendChild(RECel);
  mcform.appendChild(QTYel);

  // close the window
  document.getElementById('partEditWindow').style.visibility = 'hidden';
}
function hideEditPopUp(win){
  document.getElementById('partEditWindow').style.visibility = 'hidden';
}

function addNewNote(ver){
  if (ver==1){
    var url = "modules/workorder/mc_editnote.asp"
    top.loadmodal(url, null, 700, 500, false, false)

    if(typeof retval != "undefined" && retval !== "" && retval !== null && retval != "undefined"){
      if ($("#txtreport").val() != ""){
        $('#txtreport').val( $("#txtreport").val() + "\r" + retval); 
      } else {
        $('#txtreport').val(retval); 
      } 
    } else {}
  } else {
    top.showpopup('actions','Actions',266,100,this,txtreport)
  }

}

//RGJ END
</script>
<script language="JavaScript">
 <!--    hide
    function openOnDemandFollowupWO(btn) {
        btn.disabled = true;
        f = document.mcform;
        var param = new Object();
        param.setdefaults = true;
        param.Reason = 'Follow-up to Work Order #' + f.txtWO.value + '. (' + top.stripReturns(f.txtReason.value) + ')';
        param.FollowupFromWOPK = f.wopk.value;
        param.AssetPK = f.txtAssetPK.value;

        aW = screen.availWidth;
        aH = screen.availHeight;
        if (aW >= ((1024 * 2) - 50) || aH >= ((768 * 2) - 50)) {
            if (aW >= ((1024 * 2) - 50)) {
                aW = aW / 2;
            }
            else {
                aH = aH / 2;
            }
        }
        aW -= 50;
        aH -= 50;
        // using showModal because the workorder file is expecting the dialogArguments
        var retval = top.showModalDialog('../../default.asp?aotfmode=y&aotfmodule=WO', param, 'dialogHeight: ' + aH + 'px; dialogWidth: ' + aW + 'px; dialogTop: 25px; dialogLeft: 25px; center: No; help: No; resizable: No; status: No; scroll: No');
        
        btn.disabled = false;
        if (!retval) {            
            getNewFollowupWO(document.mcform.wopk.value, document.mcform.txtLastWOPK.value);            
        }
    }

    function getNewFollowupWO(searchvalue, searchfilter) {

        $.ajax({
            type: "POST",
            url: "workorder_close_enhanced_data.asp?action=followupwo&searchval=" + searchvalue + "&filterval=" + searchfilter + "&rc=0",
            dataType: "html",
            async: false,

            success: function(msg) {
                try {
                    //alert(msg);
                    var a = msg.split("~");

                    if (a.length == 5) {
                        if (searchfilter == 0) {
                       //     top.cleartable(ofw1);
                            document.getElementById('ofw1').style.display = '';
                        }
                        var img = '<img src="<%=Application( "web_path") & Application( "mapp_path" )%>' + a[4] + '">';
                        // alert(img);
                        document.mcform.txtLastWOPK.value = a[0];
                        top.builddatarow(ofw1body, 2, null, a[0] + '$' + a[3], a[0], 'WO', false, '', null, null, img, a[1], a[2]);
                    }
                } catch (e) { }
            }  // End of Success
        });         // End of Ajax call
    }  // End of Function getNewFollowupWO
    // JReed END

    function ofwTR_OnMouseOut(srcEle) {       
        srcEle.style.backgroundColor = "transparent";           
        srcEle.style.cursor = "";
    }
    function ofwTR_OnMouseOver(srcEle) {
        if (srcEle.style.backgroundColor.toUpperCase() == '#EFF1FA' ||
         srcEle.style.backgroundColor.toUpperCase() == '#E9EBF8' ||
         srcEle.style.backgroundColor.toUpperCase() == '#FCFDE1') {
        }
        else {
            srcEle.style.backgroundColor = "#EFEBFF";
            srcEle.style.cursor = "hand";
        }
    }
 // done hiding -->
 </script>

<link rel="stylesheet" type="text/css" href="../../css/mc_css.css" />

<style type="text/css">
#container, body.embed{
	background-color:;
}
.buttonsdisabled {
	display:	static;
	filter:		Gray() Alpha(Opacity=40);
	cursor:hand;
}
.laborassign {
    FONT-WEIGHT: normal;
    FONT-SIZE: 8pt;
}
.buttonsenabled {
	display:	;
	filter:		none;
	cursor:hand;
}
.required1
{
    BORDER-RIGHT: royalblue 1px solid;
    BORDER-TOP: royalblue 1px solid;
    PADDING-LEFT: 1px;
    FONT-WEIGHT: normal;
    FONT-SIZE: 8pt;
    MARGIN-BOTTOM: 1px;
    BORDER-LEFT: royalblue 1px solid;
    COLOR: #000000;
    BORDER-BOTTOM: royalblue 1px solid;
    FONT-FAMILY: Arial;
    BACKGROUND-COLOR: #ffffff
}
.requiredright1
{
    BORDER-RIGHT: royalblue 1px solid;
    BORDER-TOP: royalblue 1px solid;
    PADDING-LEFT: 1px;
    FONT-WEIGHT: normal;
    FONT-SIZE: 8pt;
    MARGIN-BOTTOM: 1px;
    BORDER-LEFT: royalblue 1px solid;
    COLOR: #000000;
    BORDER-BOTTOM: royalblue 1px solid;
    FONT-FAMILY: Arial;
    BACKGROUND-COLOR: #ffffff;
    TEXT-ALIGN: right
}
.popup{
  font-family: Arial;
  border-top: gray 2px ridge;
  border-left: gray 2px ridge;
  border-right: gray 2px ridge;
  border-bottom: gray 2px ridge;
  padding-top: 20px;
  padding-left: 20px;
  padding-right: 20px;
  padding-bottom: 20px;
  BACKGROUND-COLOR: white;
  position:absolute;
  margin-top: -100px;
  margin-left: -324px;
  width:500px;
  background-color:white;
  visibility:hidden;
}
.editpopup{
  font-family: Arial;
  border-top: gray 2px ridge;
  border-left: gray 2px ridge;
  border-right: gray 2px ridge;
  border-bottom: gray 2px ridge;
  padding-top: 20px;
  padding-left: 20px;
  padding-right: 20px;
  padding-bottom: 20px;
  BACKGROUND-COLOR: white;
  position:absolute;
  margin-top: -100px;
  margin-left: -324px;
  width:400px;
  background-color:white;
  visibility:hidden;
}

#list-menu {
  width: 204px;
  position: absolute;
  top:84px;
  left: 118px;
  border: 1px solid #000000;
  background: white;
  /* this width value is also effected by the padding we will later set on the links. */
}
#list-menu ul {
  margin: 0;
  padding: 0;
  list-style-type: none;
  font-family: verdana, arial, sanf-serif;
  font-size: 12px;
}
#list-menu li {
  margin: 0px 0px 0px 0px;
}
#list-menu li.sumrow {
  margin: 0px 0px 0px 0px;
  text-align: center;
  background: royalblue;
  color: #ffffff;
  font-weight: bold;
  font-family: verdana, arial, sanf-serif;
  font-size: 12px;
}
#list-menu a {
  display: block;
  width:202px;
  padding: 1px 2px 1px 7px;
  /*border: 1px solid #000000;*/
  text-decoration: none; /*lets remove the link underlines*/
}
  #list-menu a:link, #list-menu a:active, #list-menu a:visited {
  color: #000000;
}
#list-menu a:hover {
  border: 1px solid #000000;
  background: #333333;
  color: #ffffff;
}

.partrow{
  cursor: hand;
}
.partrow:hover{
  border: 1px solid #000000;
  background: #333333;
  color: #ffffff;
}

</style>
</head>

<% If Not errormessage = "" or GetSession("smallres") = "Y" Then
	  assetphoto = ""
   End If
%>

<body style="padding:15;scrollbar-base-color: #EAEAEA; font-family:Arial; font-size:8pt; color:#000000" bgColor="#FBFBFB">

<form name="mcform" id="mcform" method="post" onsubmit="return false;" target="postTo" AutoComplete="OFF">
<!-- JReed custom Constellium for FollowupWO -->
<input type="hidden" name="txtReason" value="<% =txtreason %>">
<input type="hidden" name="txtWO" value="<% =txtwo %>">
<input type="hidden" name="txtAssetPK" value="<% =txtassetpk %>">
<input type="hidden" name="txtLastWOPK" value="0">
<!-- end custom for Constellium FollowupWO -->
<input type="hidden" name="txtAccountPK" value="<% =txtaccountpk %>">
<input type="hidden" name="txtAccountDescH" value="<% =txtaccountdesc %>">
<input type="hidden" name="txtCategoryPK" value="<% =txtcategorypk %>">
<input type="hidden" name="txtCategoryDescH" value="<% =txtcategorydesc %>">
<input type="hidden" name="txtProblemPK" value="<% =txtproblempk %>">
<input type="hidden" name="txtProblemDescH" value="<% =txtproblemdesc %>">
<input type="hidden" name="txtFailurePK" value="<% =txtfailurepk %>">
<input type="hidden" name="txtFailureDescH" value="<% =txtfailuredesc %>">
<input type="hidden" name="txtSolutionPK" value="<% =txtsolutionpk %>">
<input type="hidden" name="txtSolutionDescH" value="<% =txtsolutiondesc %>">
<input type="hidden" name="mode" value="<% =mode %>">
<input type="hidden" name="wopk" value="<% =wopk %>">
<input type="hidden" name="wogrouppk" value="<% =WOGroupPK %>">
<!--RGJ BEGIN - holds current asset meter readings-->
<input type="hidden" name="assetmeter1reading" value="<%=assetmeter1%>" />
<input type="hidden" name="assetmeter2reading" value="<%=assetmeter2%>" />
<!--RGJ END -->

<%If errormessage = "" Then %>
<div style="position:absolute;left:16px;top:16px; right:16px; width:100%;height:585px;overflow:auto;">
  <table cellspacing="0" cellpadding="0" border="0" width="100%">
    <tr>
		  <td id="wono" style="font-family:Arial; font-size:12pt; color:royalblue; font-weight:bold;" align="center"><% =righttitle %></td>
	  </tr>
	  <tr>
	    <td style="<% If Not WOGroupPK = "" and Not WOGroupPK = "-1" Then %>display;<% Else %>display:none;<% End If %>cursor:hand; padding-top:4px; padding-left:0px; font-family:Arial; font-size:9pt; color:#000000;" onclick="txtWOGroupAll.click();">
		    <input onclick="this.blur();event.cancelBubble = true;top.wogrouptoggle();" type="checkbox" value="ON" name="txtWOGroupAll" tabindex="" onfocus="this.blur();">
			    All Work Orders in Group
		  </td>
    </tr>
	  <tr>
	    <td bgcolor="#DBB72E" height="1" style="background-color:#FFCC00;font-family: arial; font-size: 9pt; color: #000000"><img src="../../images/blank.gif" border="0" width="5" height="1" /></td>
	  </tr>
  </table>

  <table id="woContent" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; padding-top:6px;" bordercolor="#111111" width="100%">
    <tr>
      <td valign="top" width="100%">
        <table cellspacing="0" cellpadding="0" border="0" width="100%">
        <!-- RGJ Build the page -->
        <% Call BuildContent() %>
        
        </table>
	    </td>
    </tr>
  </table>
</div>

  <%
  writerecorddata
  %>

  <table style="position:absolute; top:610px; left:20px; font-family:Arial; font-size:9pt; color:#000000" border="0" cellpadding="0" cellspacing="0" width="100%">
	  <tr>
		  <td>
			  <span style="display:none; cursor:hand;" onclick="txtDefault.click();">
  			  <input onclick="this.blur();" type="checkbox" value="ON" name="txtDefault" tabindex="">Set as Default
			  </span>
		  </td>
		  <td align="right" style="padding-right:13px;">
		    <%
			  If Not WO_CLOSE_CUSTOMHOOK = "" Then
				  Response.Write(Replace(WO_CLOSE_CUSTOMHOOK,"##ASSETPK##",AssetPK)) & "&nbsp;&nbsp;&nbsp;&nbsp;"
			  End If
		    %>
			  <img name="applybutton" id="applybutton" border="0" src="../../images/buttonaction_apply.gif" style="cursor:hand;" width="80" height="19" onclick="doChecks();" alt="" />
			  <img name="applybutton" id="cancelbutton" border="0" src="../../images/buttonaction_cancel.gif" style="cursor:hand;" width="80" height="19" onclick="top.applyit = false; top.close();" alt="" />
		  </td>

	  </tr>

    <%If WO_CLOSE_ALLTASKSCOMPLETE_REQ = "Yes" Then%>
	  <tr><td colspan="2" align="right"><font size="2"><span id="taskMSG" name="taskMSG" style="padding-right:30px;<%If TaskCode = -1 Then Response.write "color:red;font-weight:bold;" Else If TaskCode = 1 Then Response.Write "color:green;" Else Response.Write "display:none;" End If%>"><%=TaskMessage%></span></font></td></tr>
    <%End If%>

    <%
    'If NullCheck(assetphoto) <> "" Then 
    '  Dim zoomimage
    '  zoomimage = Replace(assetphoto,"_wo.",".", 1, -1, vbTextCompare)
    '  'Response.write zoomimage
    %>
    <!--
        <tr>
          <td colspan="2">
            <div id="assetphotobox" style="display:; margin-top:0; text-align: left; vertical-align: bottom; background-color: transparent; width:150; height:110; overflow-y:auto; overflow-x:auto;">
              <img onclick="top.openimgzoom('<%'=zoomimage%>'); event.returnValue = false;" style="cursor:hand;" id="assetphoto" border="0" src="<%'=assetphoto %>" align="left" alt="" />
            </div>      
          </td>
        </tr>
    -->
    <%'End If%>
  </table>

<% Else %>
  <div style="font-family:Arial; font-size:10pt; color:#000000;">
		<% =errormessage %>
	</div>
	<table <% 'position:absolute; top:600px; left:20px; %>style="position:absolute; top:650px; left:20px; font-family:Arial; font-size:9pt; color:#000000" border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td></td>
			<td align="right" style="padding-right:13px;">
			  <img id="cancelbutton" onclick="top.applyit = false; top.close();" border="0" src="../../images/buttonaction_cancel.gif" style="cursor:hand;" WIDTH="80" HEIGHT="19">
			</td>
		</tr>
	</table>
<% End If %>


<div id="loadingdiv" style="position:absolute;display:none; background-Color:white; border:1px solid grey; border-top:2px solid #cccccc; border-left:2px solid #cccccc; height:15px; width:20px;z-index:100;">
	<table bgcolor="#FFFFFF" border="0" cellpadding="0" cellspacing="0" width="100%" height="100%"><tr><td width="100%" align="center"><font size="1" face="Arial"><center>Loading...Please Wait</center></font></td></tr></table>
</div>

<div id="lookupsavecancel" STYLE="display: none; position:absolute; z-index:100;">
	<table bgcolor="#FFFFCC" border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
	<td width="100%" align="center" height="25" valign="bottom" background="../../images/lookuptopbg.jpg">
	<img border="0" style="cursor:hand;" src="../../images/button_save.gif" width="80" height="15" onclick="document.frapopupouter.submitform('SAVE','Saving');event.cancelBubble = true;">&nbsp;<img style="cursor:hand;" border="0" src="../../images/button_cancel.gif" width="80" height="15" onclick="document.frapopupouter.cancelform();event.cancelBubble = true;">
	</td>
	</tr>
	</table>
</div>

<div id="lookuppopupupper" STYLE="display: none; position:absolute; z-index:100;">
	<table border="0" width="100%" height="20" bgcolor="#FFFFCC" CELLSPACING="0" CELLPADDING="0">
	<tr>
	  <th width="100%" NOWRAP align="left">
	    <table border="0" cellpadding="0" cellspacing="0" width="100%">
	      <tr>
	        <td>
	        <table CELLSPACING="0" CELLPADDING="0" border="0"><tr><td><img border="0" src="../../images/red-arrow.gif" width="8" height="12" ondblclick="document.frapopupouter.dolookupedit();event.cancelBubble = true;"></td><td><font style="font-family: Arial; font-size: 8pt; font-weight: bold" color="#000000">&nbsp;<span id="lookuppopupuppertext">Lookup Table</span></font></td></tr></table></td>
	        <td align="right"><font style="font-family: Arial; font-size: 7pt; font-weight: bold; color=#FFFFFF" color="#FFFFFF"><img id="lookuptableeditbutton" border="0" style="cursor:hand;" src="../../images/popupedit.gif" width="26" height="14" onclick="document.frapopupouter.dolookupedit();event.cancelBubble = true;"><img border="0" style="cursor:hand;" src="../../images/closepopup.gif" hspace="1" width="15" height="14" onclick="top.clearallpopups();event.cancelBubble = true;" title="Close"></font></td>
	      </tr>
	    </table>
	  </th>
	</tr>
	</table>
</div>

<div id="caption_l" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 267; height: 100; padding-left: 28; padding-right: 12; padding-top: 10; background-image: url('../../images/caption_l.gif'); z-index:100;"></div>
<div id="caption_r" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 267; height: 100; padding-left: 12; padding-right: 28; padding-top: 10; background-image: url('../../images/caption_r.gif'); z-index:100;"></div>
<div id="caption_tl" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 28; background-image: url('../../images/caption_tl.gif'); z-index:100;"></div>
<div id="caption_tr" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 28; background-image: url('../../images/caption_tr.gif'); z-index:100;"></div>
<div id="caption_tr_2" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 28; background-image: url('../../images/caption_tr2.gif'); z-index:100;"></div>
<div id="caption_bl" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 10; background-image: url('../../images/caption_bl.gif'); z-index:100;"></div>
<div id="caption_br" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 10; background-image: url('../../images/caption_br.gif'); z-index:100;"></div>

<iframe height="0" width="0" style="display:none;position:absolute;z-index=100;" id="frapopupouter" name="frapopupouter" MARGINHEIGHT="0" MARGINWIDTH="0" NORESIZE FRAMEBORDER="0" SCROLLING="yes" allowTransparency="true" SRC="../../mc_sslfix.htm"></iframe>
<iframe height="0" width="0" style="display:; position:absolute; top:600px;" id="postTo" name="postTo" MARGINHEIGHT="0" MARGINWIDTH="0" NORESIZE FRAMEBORDER="0" SCROLLING="yes" allowTransparency="true" SRC="../../mc_sslfix.htm"></iframe>
</form>
</body>

</html>

<%
Sub checkforsubmit()
	If Request.Form("mode") = "" Then
		Exit Sub
	End If

	' Handle the Assignments
	Dim completeassignments
	If Request.Form("completeassignments") = "" Then
		completeassignments	= 0
	Else
		completeassignments = 1
	End If

	' Reset Mode to the one in the Form collection (versus the one in the Query String)
	Mode = Trim(Request.Form("Mode"))

	responded = UCase(Request.Form("txtResponded"))
	If responded = "Y" Then
		responded = True
	Else
		responded = False
	End If
	completed = UCase(Request.Form("txtCompleted"))
	If completed = "Y" Then
		completed = True
	Else
		completed = False
	End If
	finalized = UCase(Request.Form("txtFinalized"))
	If finalized = "Y" Then
		finalized = True
	Else
		finalized = False
	End If
	closed = UCase(Request.Form("txtClosed"))
	If closed = "Y" Then
		closed = True
	Else
		closed = False
	End If

	%>
	<html>
	<head>
	<script language="javascript">

		var completed = false;

		var param = new Object();
		param.caller = self;
		var p = null;

		// for debugging
		if (parent == null)
		{
			p = self.opener;
		}
		else
		{
			p = parent;
		}

		// to debug set to false and change the form target to a window
		// that does not exist
		<% If Not mode = "WO" Then %>
		if (true)
		{
			if (p.processwindow && p.processwindow.closed == false)
			{
				p.processwindow.focus();
			}
			else
			{
				p.processwindow = p.showModelessDialog('workorder_close_message.htm', param,'dialogHeight: 160px; dialogWidth: 400px; dialogTop: px; dialogLeft: px; center: Yes; help: No; resizable: No; status: No; scroll: no' );
			}
		}
		<% End If %>
		//parent.generateform.style.display = 'none';
		//parent.buttontable.style.display = 'none';
		//parent.processdiv.style.display = '';
		//parent.close();

		function doonload()
		{
			if (p.processwindow && p.processwindow.closed == false)
			{
				p.processwindow.mcworking.style.display = 'none';
				p.processwindow.mcdone.style.display = '';
			}
			completed = true;

			p.playsound('sounds/done.wav');
			<% If mode = "WO" Then %>
			try {
				if (p.dialogArguments.smartwindow != null)
				{
					if (p.dialogArguments.smartwindow.updatesmartrow)
					{
						p.dialogArguments.smartwindow.updatesmartrow('<% =Responded %>','<% =Completed %>','<% =Finalized %>','<% =Closed %>');
					}
				}
			} catch(e) {}
			doclose();
			<% End If %>
		}

		function doclose()
		{
			p.applyit = true;
			<% If UCase(Request.QueryString("fromcal")) = "Y" Then %>
			var calwin = p.dialogArguments.calwindow;
			if (calwin != null)
			{
				calwin.location.replace(calwin.location.href);
			}
			p.afterpost();
			<% ElseIf curmod = "WR" Then %>
			p.afterpost('top.fraTopic.fraPreview.refreshreport();top.refreshcurrentexplorer();');
			<% Else %>
			<% If mode = "WO" Then %>
			p.afterpost('top.refreshcurrentrecord();top.refreshcurrentexplorer(true);');
			<% Else %>
			if (p.thetop.splitview == true || p.thetop.maxview == true)
			{
				// for the Group Tab of WO or WO tab of Project
				p.afterpost('top.refreshcurrentrecord();top.refreshcurrentexplorer(true);');
			}
			else
			{
				p.afterpost('top.refreshcurrentexplorer();');
			}
			<% End If %>
			<% End If %>
			p.close();
		}

	</script>
	</head>
	<body onload="doonload();">
	<%
	Call FlushIt()

	'Call aspdebug
	'Response.End

  Dim txtRESPONDEDOVERWRITE, txtCOMPLETEDOVERWRITE, txtFINALIZEDOVERWRITE

  txtRESPONDEDOVERWRITE = False
  txtCOMPLETEDOVERWRITE = False
  txtFINALIZEDOVERWRITE = False

'RGJ BEGIN
'Labor Actuals added by user
'Response.Write mode
'Response.End
'Dim q
'response.write "<b>Form Elements: " &"</b><br>"
'for q = 1 to Request.Form.Count 
'  Response.Write q & "  - " & Request.Form.key(q) & " - " & request(Request.Form.key(q)) & "<br>"
'next
'response.write "<br><b>Querystring Elemenets: " &"</b><br>"
'for q = 1 to Request.QueryString.Count 
'  Response.Write q & "  - " & Request.QueryString.key(q) & " - " & request(Request.QueryString.key(q)) & "<br>"
'next 
'Response.End

'Clear db error object
db.dok = True
db.derror = ""

'Get User Info
Dim rowuser, rowinitials, rowuserip
rowuser = GetSession("USERPK")
rowinitials = GetSession("USERINITIALS")
rowuserip = GetSession("USERIPADDRESS")

Dim lpk,wd,reg,ovt,oth,lcom(),ci,item,X,SplitLabor,workdate
SplitLabor = 0
If Request.Form("disperseLabor") = "on" Then
  SplitLabor = 1
End if

If NullCheck(Request.Form("LaborPK")) <> "" Then
  if instr(Request.Form("LaborPK"),",") > 0 then
	  lpk = split(Request.Form("LaborPK"),",")
	  wd = split(Request.Form("WorkDate"),",")
	  reg = split(Request.Form("RegularHours"),",")
	  ovt = split(Request.Form("OvertimeHours"),",")
	  oth = split(Request.Form("OtherHours"),",")
	  ReDim lcom(Request.Form("LaborComments").Count)
	  ci=0
	  For Each item In Request.Form("LaborComments")
		  lcom(ci)=item
		  ci=ci+1
		Next
  else
	  lpk = array(Request.Form("LaborPK"))
	  wd = array(Request.Form("WorkDate"))
	  reg = array(Request.Form("RegularHours"))
	  ovt = array(Request.Form("OvertimeHours"))
	  oth = array(Request.Form("OtherHours"))
	  ReDim lcom(1)
	  lcom(0) = Request.Form("LaborComments")
  end if
End If

If NullCheck(Request.Form("LaborPK")) <> "" And mode = "WO" Then  
    For X = (LBound(lpk)) To (UBound(lpk))
      workdate = SQLdatetime(wd(X))

      'Check to see if Labor Comments is blank - this was added to fix issue with Labor Time Sheet and a Distinct query problem
      If NullCheck(Trim(lcom(X))) <> "" Then
        sql = "INSERT INTO WOLabor (WOPK, LaborPK, LaborID, LaborName, RecordType, LaborType, laborTypeDesc, WorkDate, RegularHours, OvertimeHours, OtherHours, AutoCalcCost, CostRegular, CostOvertime, CostOther, ChargeRate, ChargePercentage, Comments, RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionDate) "+_
        "SELECT "&WOPK&", LaborPK, LaborID, LaborName, 2, LaborType, LaborTypeDesc, '"&workdate&"',"&Round(reg(X),2)&","&Round(ovt(X),2)&","&Round(oth(X),2)&",1, CostRegular, CostOvertime, CostOther, ChargeRate, ChargePercentage, '"&SQLEncode(lcom(X))&"', '"&rowuserip&"', "&rowuser&", '"&rowinitials&"', GETDATE() FROM Labor WHERE LaborPK = "&lpk(X)
      Else
        sql = "INSERT INTO WOLabor (WOPK, LaborPK, LaborID, LaborName, RecordType, LaborType, laborTypeDesc, WorkDate, RegularHours, OvertimeHours, OtherHours, AutoCalcCost, CostRegular, CostOvertime, CostOther, ChargeRate, ChargePercentage, RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionDate) "+_
        "SELECT "&WOPK&", LaborPK, LaborID, LaborName, 2, LaborType, LaborTypeDesc, '"&workdate&"',"&Round(reg(X),2)&","&Round(ovt(X),2)&","&Round(oth(X),2)&",1, CostRegular, CostOvertime, CostOther, ChargeRate, ChargePercentage, '"&rowuserip&"', "&rowuser&", '"&rowinitials&"', GETDATE() FROM Labor WHERE LaborPK = "&lpk(X)
      End If
      'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
      'Response.End
      Call db.RunSQL(sql,"")
      Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
    Next
End If

'Part Actuals added by User

Dim plpk, ppk, qty, apk, aid, anm, cpk, cid, cnm, vpk, vid, vnm, act, cat, expd, ord, delv, comm
  
If NullCheck(Request.Form("PartPK")) <> "" Then
  if instr(Request.Form("PartPK"),",") > 0 then
	  plpk = split(Request.Form("LocationPK"),",")
	  ppk = split(Request.Form("PartPK"),",")
	  qty = split(Request.Form("Quantity"),",")
	  act = split(Request.Form("AccountID"),",")
	  cat = split(Request.Form("CategoryID"),",")
	  delv = split(Request.Form("UDFChar1"),",")
	  expd = split(Request.Form("UDFBit2"),",")
	  ord = split(Request.Form("UDFBit1"), ",")
	  comm = split(Request.Form("WOP_Comments"),",")
  else
	  plpk = array(Request.Form("LocationPK"))
	  ppk = array(Request.Form("PartPK"))
	  qty = array(Request.Form("Quantity"))
	  act = array(Request.Form("AccountID"))
	  cat = array(Request.Form("CategoryID"))
	  delv = array(Request.Form("UDFChar1"))
	  expd = array(Request.Form("UDFBit2"))
	  ord = array(Request.Form("UDFBit1"))
	  comm = array(Request.Form("WOP_Comments"))
  end if
End If

If NullCheck(Request.Form("PartPK")) <> "" And mode = "WO" Then  
    For X = (LBound(ppk)) To (UBound(ppk))
      'Get Account Info
      'If IsArray(aid) Then 
        sql = "SELECT AccountPK, AccountID, AccountName FROM Account WHERE AccountID = '"&Trim(act(X))&"'"
        'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
        'Response.End
        Set rs = db.RunSqlReturnRS(sql,"")
        If not rs.EOF Then
          apk = NullCheck(rs("AccountPK"))
          aid = NullCheck(rs("AccountID"))
          anm = NullCheck(rs("AccountName"))
        Else
          apk = ""
          aid = ""
          anm = ""
        End If
      'Else
      '  apk=""
      '  aid=""
      '  anm=""
      'End If

      'Get Category Info
      'If IsArray(cid) Then
        sql = "SELECT CategoryPK, CategoryID, CategoryName FROM Category WHERE CategoryID = '"&Trim(cat(X))&"'"
        'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
        'Response.End
        Set rs = db.RunSqlReturnRS(sql,"")
        If not rs.EOF Then
          cpk = NullCheck(rs("CategoryPK"))
          cid = NullCheck(rs("CategoryID"))
          cnm = NullCheck(rs("CategoryName"))
        Else
          cpk = ""
          cid = ""
          cnm = ""
        End If
      'Else
      '  cpk = ""
      '  cid = ""
      '  cnm = ""
      'End If
      'Response.Write "plpk(X): " & plpk(X) & "<br>"
      'Response.Write "aid(X): " & aid(X) & "<br>"
      Dim isDI, isOOP
      isDI=0
      isOOP=0 
      
      If NullCheck(Trim(plpk(X))) = "99999" Then
        isDI=1
        isOOP=0
      Else
        isDI=0
        isOOP=1
      End If
    
      If (Trim(plpk(X)) = "-1" Or Trim(plpk(X)) = "99998" Or Trim(plpk(X)) = "99999") Then
        sql = "INSERT INTO WOPart (WOPK, RecordType, PartPK, PartID, PartName, DirectIssue, OutOfPocket, QuantityEstimated, IssueUnitCost, IssueUnitChargePrice," +_
        " TotalCost, AutoCalcCost, AccountPK, AccountID, AccountName, CategoryPK, CategoryID, CategoryName, UDFChar1, UDFBit1, UDFBit2, Comments," +_
        " RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionDate) "+_
        "SELECT "&WOPK&", 1, Part.PartPK, Part.PartID, Part.PartName, " & isDI & ", " & isOOP & ", "
        sql = sql & qty(X)&",Part.IssueUnitCost,Part.IssueUnitChargePrice,Part.IssueUnitCost*"&qty(X)&",1, "
        If apk = "" Or apk = "-1" Then
          sql = sql & "NULL, NULL, NULL, "
        Else
          sql = sql & apk & ",'" & aid & "','" & anm & "', "
        End If
        If cpk = "" Or cpk = "-1" Then
          sql = sql & "NULL, NULL, NULL "
        Else
          sql = sql & cpk & ",'" & cid & "','" & cnm & "' "
        End If
        
        if not IsArray(delv) or X > Ubound(delv) then
            sql = sql & ", NULL"
        else
            sql = sql & ", '" &Replace(Trim(delv(X)), "'","''")& "' " 
        end if
        if not IsArray(ord) or X > Ubound(ord) then
            sql = sql & ", NULL"
        else
            sql = sql & "," &ord(X)
        end if
        if not IsArray(expd) or X > Ubound(expd) then
            sql = sql & ", NULL"
        else
            sql = sql & "," &expd(X)
        end if
         if not IsArray(comm) or X > Ubound(comm) then
            sql = sql & ", NULL"
        else
            sql = sql & ", '" &Replace(Trim(comm(X)), "'","''")& "' " 
        end if        
                
        sql = sql & ", '"&rowuserip&"', "&rowuser&", '"&rowinitials&"', GETDATE() FROM Part WITH (NOLOCK) WHERE PartPK = "&ppk(X)
      Else
        sql = "INSERT INTO WOPart (WOPK, RecordType, PartPK, PartID, PartName, LocationPK, LocationID, LocationName, QuantityEstimated, IssueUnitCost, IssueUnitChargePrice," +_
        " TotalCost, AutoCalcCost, AccountPK, AccountID, AccountName, CategoryPK, CategoryID, CategoryName, UDFChar1, UDFBit1, UDFBit2, Comments," +_
        " RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionDate) "+_
        "SELECT "&WOPK&", 1, Part.PartPK, Part.PartID, Part.PartName, Location.LocationPK, Location.LocationID, Location.LocationName, "&qty(X)&",PartLocation.IssueUnitCost,PartLocation.IssueUnitChargePrice,PartLocation.IssueUnitCost*"&qty(X)&",1, "
        If apk = "" Or apk = "-1" Then
          sql = sql & "NULL, NULL, NULL, "
        Else
          sql = sql & apk & ",'" & aid & "','" & anm & "', "
        End If
        If cpk = "" Or cpk = "-1" Then
          sql = sql & "NULL, NULL, NULL "
        Else
          sql = sql & cpk & ",'" & cid & "','" & cnm & "' "
        End If
        if not IsArray(delv) or X > Ubound(delv) then
            sql = sql & ", NULL"
        else
            sql = sql & ", '" &Replace(Trim(delv(X)), "'","''")& "' " 
        end if
        if not IsArray(ord) or X > Ubound(ord) then
            sql = sql & ", NULL"
        else
            sql = sql & "," &ord(X)
        end if
        if not IsArray(expd) or X > Ubound(expd) then
            sql = sql & ", NULL"
        else
            sql = sql & "," &expd(X)
        end if
         if not IsArray(comm) or X > Ubound(comm) then
            sql = sql & ", NULL"
        else
            sql = sql & ", '" &Replace(Trim(comm(X)), "'","''")& "' " 
        end if      
        
        sql = sql & ", '"&rowuserip&"', "&rowuser&", '"&rowinitials&"', GETDATE() FROM Part WITH (NOLOCK) INNER JOIN PartLocation ON PartLocation.PartPK = Part.PartPK " +_
        "INNER JOIN Location ON Location.LocationPK = PartLocation.LocationPK WHERE PartLocation.PartPK = "&ppk(X)&" AND PartLocation.LocationPK = "&plpk(X)
      End If
      'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea><br>"
      'Response.End
      Call db.RunSQL(sql,"")
      Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
    Next
End If

'Part QTY Updates
If mode = "WO" Then
  Dim woppk, npqty
  If NullCheck(Request.Form("WOPartPK")) <> "" Then
    if instr(Request.Form("WOPartPK"),",") > 0 then
	    woppk = split(Request.Form("WOPartPK"),",")
	    npqty = split(Request.Form("UPDPartQTY"),",")
    else
	    woppk = array(Request.Form("WOPartPK"))
	    npqty = array(Request.Form("UPDPartQTY"))
    end if

    For X = (LBound(woppk)) To (UBound(woppk))
      Dim xppk, xpk, rvd
      xppk = split(woppk(x),"$")
      xpk = xppk(0)
      rvd = xppk(1)
      sql = "UPDATE WOPart SET QuantityEstimated = " & npqty(x) & " WHERE PK = " & xpk
      'response.Write "<textarea>"&sql&"</textarea>"
      'Response.End
      Call db.RunSQL(sql,"")
      Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
    Next
  End If
End If


'Misc Cost Actuals

Dim mcn, mcd(), mcdate,mcv,mca, mcc, mcp, mcom()
If NullCheck(Request.Form("txtMCName")) <> "" Then

  if instr(Request.Form("txtMCName"),",") > 0 then
	  mcn = split(Request.Form("txtMCName"),",")
	  ReDim mcd(Request.Form("txtMCDesc").Count)
	  ci=0
	  For Each item In Request.Form("txtMCDesc")
		  mcd(ci)=item
		  ci=ci+1
		Next
	  'mcd = split(Request.Form("txtMCDesc"),",")
	  mcdate = split(Request.Form("txtMCDate"),",")
	  mcv = split(Request.Form("txtMCVendorID"),",")
	  mca = split(Request.Form("txtMCAccountID"),",")
	  mcc = split(Request.Form("txtMCCategoryID"),",")
	  mcp = split(Request.Form("txtMCPrice"),",")
	  ReDim mcom(Request.Form("MiscCostComments").Count)
	  ci=0
	  For Each item In Request.Form("MiscCostComments")
		  mcom(ci)=item
		  ci=ci+1
		Next
	  'mcom = split(Request.Form("MiscCostComments"),",")
  else
	  mcn = array(Request.Form("txtMCName"))
	  ReDim mcd(1)
	  mcd(0) = Request.Form("txtMCDesc")
	  mcdate = array(Request.Form("txtMCDate"))
	  mcv = array(Request.Form("txtMCVendorID"))
	  mca = array(Request.Form("txtMCAccountID"))
	  mcc = array(Request.Form("txtMCCategoryID"))
	  mcp = array(Request.Form("txtMCPrice"))
	  ReDim mcom(1)
	  mcom(0) = Request.Form("MiscCostComments")
  end if
End If

If NullCheck(Request.Form("txtMCName")) <> "" And mode = "WO" Then 
    For X = (LBound(mcn)) To (UBound(mcn))
      Dim MCPrice
      MCPrice=0

      'Get Vendor Info
      If IsArray(mcv) Then
        'If mcv(X) <> "-1"  Then
        sql = "SELECT CompanyPK, CompanyID, CompanyName FROM Company WHERE CompanyID = '"&Trim(mcv(X))&"'"
        'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
        'Response.End
        Set rs = db.RunSqlReturnRS(sql,"")
        If Not db.dok then
          Response.Write "Vendor Error"
          Response.End
        Else
          If not rs.EOF Then
            vpk = NullCheck(rs("CompanyPK"))
            vid = NullCheck(rs("CompanyID"))
            vnm = NullCheck(rs("CompanyName"))
          Else
            vpk=""
            vid=""
            vnm=""
          End If
        End If
        'End if
      Else
        vpk=""
        vid=""
        vnm=""
      End If

      'Get Account Info
      If IsArray(mca) Then
        'If mca(X) <> "-1"  Then
        sql = "SELECT AccountPK, AccountID, AccountName FROM Account WHERE AccountID = '"&Trim(mca(X))&"'"
        'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
        'Response.End
        Set rs = db.RunSqlReturnRS(sql,"")
        If Not db.dok then
          Response.Write "Misc Cost Account Error"
          Response.End
        Else
          If not rs.EOF Then
            apk = NullCheck(rs("AccountPK"))
            aid = NullCheck(rs("AccountID"))
            anm = NullCheck(rs("AccountName"))
          End If
        End If
        'End if
      Else
        apk = NullCheck(rs("AccountPK"))
        aid = NullCheck(rs("AccountID"))
        anm = NullCheck(rs("AccountName"))
      End If

      If IsArray(mcc) Then
        'If mcc(X) <> "-1" Then
        'Get Category Info
        sql = "SELECT CategoryPK, CategoryID, CategoryName FROM Category WHERE CategoryID = '"&Trim(mcc(X))&"'"
        'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
        'Response.End
        Set rs = db.RunSqlReturnRS(sql,"")
        If Not db.dok then
          Response.Write "Misc Cost Account Error"
          Response.End
        Else
          If not rs.EOF Then
            cpk = NullCheck(rs("CategoryPK"))
            cid = NullCheck(rs("CategoryID"))
            cnm = NullCheck(rs("CategoryName"))
          Else
            cpk=""
            cid=""
            cnm=""
          End If
        End If
        'End If
      Else
        cpk=""
        cid=""
        cnm=""
      End If
      '20100823 - RGJ Added fix to Total Cost - if null then zero
      If IsArray(mcp) Then
        If Trim(NullCheck(mcp(X))) = "" Then
          MCPrice = 0
        Else
          MCPrice = Trim(NullCheck(mcp(X)))
        End If
      Else
        MCPrice = 0
      End If

      '20091120 - RGJ - Added SQLDateTime on insert to ensure proper funcationality regardless of date format
      workdate = SQLdatetime(mcdate(X))
      sql = "INSERT INTO WOMiscCost (WOPK, MiscCostName, MiscCostDesc, MiscCostDate, RecordType, EstimatedCost, AccountPK, AccountID, AccountName, CategoryPK, CategoryID, CategoryName, CompanyPK, CompanyID, CompanyName, Comments, RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionDate) "+_
      "VALUES ("&WOPK&", '"&SQLEncode(mcn(X))&"', '"&SQLEncode(mcd(X))&"', '"&workdate&"', 1, "&SQLEncode(MCPrice)&","
      If apk = "" Or apk = "-1" Then
        sql = sql & "NULL, NULL, NULL, "
      Else
        sql = sql & apk & ",'" & aid & "','" & SQLEncode(anm) & "', "
      End If
      If cpk = "" Or cpk = "-1" Then
        sql = sql & "NULL, NULL, NULL,"
      Else
        sql = sql & cpk & ",'" & cid & "','" & SQLEncode(cnm) & "',"
      End If
      If vpk = "" Or vpk = "-1" Then
        sql = sql & "NULL, NULL, NULL,"
      Else
        sql = sql & vpk & ",'" & vid & "','" & SQLEncode(vnm) & "',"
      End If
      sql = sql & "'"&SQLEncode(mcom(X))&"', '"&rowuserip&"', "&rowuser&", '"&rowinitials&"', GETDATE())"
      'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
      'Response.End

      Call db.RunSQL(sql,"")
      Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

    Next
End If

'RGJ END

	requesteddate = DateNullCheck(Request.Form("txtrequestedDate"))
	requestedtime = TimeNullCheck(Request.Form("txtrequestedTime"))
	requestedinitials = Trim(Request.Form("txtrequestedinitials"))
	issueddate = DateNullCheck(Request.Form("txtissuedDate"))
	issuedtime = TimeNullCheck(Request.Form("txtissuedTime"))
	issuedinitials = Trim(Request.Form("txtissuedinitials"))
	respondeddate = DateNullCheck(Request.Form("txtrespondedDate"))
	respondedtime = TimeNullCheck(Request.Form("txtrespondedTime"))
	respondedinitials = Trim(Request.Form("txtrespondedinitials"))
	If Not Request.Form("txtRESPONDEDOVERWRITE") = "" Then
	    txtRESPONDEDOVERWRITE = True
	End If
	completeddate = DateNullCheck(Request.Form("txtCompletedDate"))
	completedtime = TimeNullCheck(Request.Form("txtCompletedTime"))
	completedinitials = Trim(Request.Form("txtcompletedinitials"))
	If Not Request.Form("txtCOMPLETEDOVERWRITE") = "" Then
	    txtCOMPLETEDOVERWRITE = True
	End If
	finalizeddate = DateNullCheck(Request.Form("txtfinalizedDate"))
	finalizedtime = TimeNullCheck(Request.Form("txtfinalizedTime"))
	finalizedinitials = Trim(Request.Form("txtfinalizedinitials"))
	If Not Request.Form("txtFINALIZEDOVERWRITE") = "" Then
	    txtFINALIZEDOVERWRITE = True
	End If
	closeddate = DateNullCheck(Request.Form("txtclosedDate"))
	closedtime = TimeNullCheck(Request.Form("txtclosedTime"))
	closedinitials = Trim(Request.Form("txtclosedinitials"))

	laborreport = Trim(Mid(Replace(Request.Form("txtreport"),Chr(13)+Chr(10),"%0D%0A"),1,6000))

	lri = Trim(Request.Form("lri"))	' Labor Report Options 1-Do Nothing 2-Append 3-Overwrite
	If Not lri = "" Then
		lri = CInt(lri)
	Else
		lri = 3
	End If
	If mode = "WO" Then
		lri = 3
	End If

	txtaccountpk = Trim(Request.Form("txtaccountpk"))
	If txtaccountpk = "" Then
		txtaccountpk = Null
	Else
		txtaccountpk = CLng(txtaccountpk)
	End If


	txtaccount = Trim(Request.Form("txtaccount"))
	txtaccountdesc = Trim(Request.Form("txtaccountdesch"))
	txtAccountAll = Trim(Request.Form("txtAccountAll"))
	If Not txtAccountAll = "" Then
		txtAccountAll = True
	Else
		txtAccountAll = False
	End If
	txtchargeable = Trim(Request.Form("txtchargeable"))
	If Not txtchargeable = "" Then
		txtchargeable = True
	Else
		txtchargeable = False
	End If
	txtcategorypk = Trim(Request.Form("txtcategorypk"))
	If txtcategorypk = "" Then
		txtcategorypk = Null
	Else
		txtcategorypk = CLng(txtcategorypk)
	End If
	txtcategory = Trim(Request.Form("txtcategory"))
	txtcategorydesc = Trim(Request.Form("txtcategorydesch"))
	txtCategoryAll = Trim(Request.Form("txtCategoryAll"))
	If Not txtCategoryAll = "" Then
		txtCategoryAll = True
	Else
		txtCategoryAll = False
	End If
	txtTasks = Trim(Request.Form("txtTasks"))
	If Not txtTasks = "" Then
		txtTasks = True
	Else
		txtTasks = False
	End If
	txtTaskInitials = Trim(Request.Form("txtTaskInitials"))
	txtLabor1 = Trim(Request.Form("txtLabor1"))
	If Not txtLabor1 = "" Then
		txtLabor1 = True
	Else
		txtLabor1 = False
	End If
	txtLabor3 = Trim(Request.Form("txtLabor3"))
	If Not txtLabor3 = "" Then
		txtLabor3 = True
	Else
		txtLabor3 = False
	End If
	txtMyLaborHrs = Trim(Request.Form("txtMyLaborHrs"))
	If txtMyLaborHrs = "" Then
	    txtMyLaborHrs = Null
	End If
	txtMaterials = Trim(Request.Form("txtMaterials"))
	If Not txtMaterials = "" Then
		txtMaterials = True
	Else
		txtMaterials = False
	End If
	txtOtherCost = Trim(Request.Form("txtOtherCost"))
	If Not txtOtherCost = "" Then
		txtOtherCost = True
	Else
		txtOtherCost = False
	End If
	txtproblempk = Trim(Request.Form("txtproblempk"))
	If txtproblempk = "" Then
		txtproblempk = Null
	Else
		txtproblempk = CLng(txtproblempk)
	End If
	txtproblem = Trim(Request.Form("txtproblem"))
	txtproblemdesc = Trim(Request.Form("txtproblemdesch"))
	txtfailurepk = Trim(Request.Form("txtfailurepk"))
	If txtfailurepk = "" Then
		txtfailurepk = Null
	Else
		txtfailurepk = CLng(txtFailurepk)
	End If
	txtfailure = Trim(Request.Form("txtfailure"))
	txtfailuredesc = Trim(Request.Form("txtfailuredesch"))
	txtsolutionpk = Trim(Request.Form("txtsolutionpk"))
	If txtsolutionpk = "" Then
		txtsolutionpk = Null
	Else
		txtsolutionpk = CLng(txtsolutionpk)
	End If
	txtsolution = Trim(Request.Form("txtsolution"))
	txtsolutiondesc = Trim(Request.Form("txtsolutiondesch"))
	txtfailurewo = Trim(Request.Form("txtFailedWO"))
	If Not txtfailurewo = "" Then
		txtfailurewo = True
	Else
		txtfailurewo = False
	End If
	If mode = "WO" Then
		txtmeter1reading = Trim(Request.Form("txtmeter1reading"))
		If txtmeter1reading = "" Then
			txtmeter1reading = 0
		End If
		txtmeter2reading = Trim(Request.Form("txtmeter2reading"))
		If txtmeter2reading = "" Then
			txtmeter2reading = 0
		End If
	Else
		txtmeter1reading = 0
		txtmeter2reading = 0
	End If
	txtisup = Trim(Request.Form("txtDownTime"))
	If Not txtisup = "" Then
		txtisup = True
	Else
		txtisup = False
	End If
	txtDrawingUpdatesNeeded = Trim(Request.Form("txtDrawingUpdatesNeeded"))
	If Not txtDrawingUpdatesNeeded = "" Then
		txtDrawingUpdatesNeeded = True
	Else
		txtDrawingUpdatesNeeded = False
	End If
	'Response.Write txtsolutiondesc
	'Response.End

	Dim txtMyLaborHrsPK
	Dim txtMyLaborOHrs
  Dim txtFollowUpSingleWO
  Dim txtFollowUpMultiWO

  txtMyLaborHrsPK = Null
  txtMyLaborOHrs = Null
  txtFollowUpSingleWO = Trim(Request.Form("txtFollowUpSingleWO"))
	If Not txtFollowUpSingleWO = "" Then
		txtFollowUpSingleWO = True
	Else
		txtFollowUpSingleWO = False
	End If
	txtFollowUpMultiWO = Trim(Request.Form("txtFollowUpMultiWO"))
	If Not txtFollowUpMultiWO = "" Then
		txtFollowUpMultiWO = True
	Else
		txtFollowUpMultiWO = False
	End If


	If mode = "WO" Then
		' ************************************************************************************************************
		' CLOSE SINGLE WORK ORDER
		' ************************************************************************************************************
    'Update the WO cost fields
    Call db.RunSP("MC_CalcWorkOrder",Array(Array("@WOPK", adInteger, adParamInput, 4, WOPK)),"")

		If completeassignments = 1 Then
			sql = "Update WOAssignStatus SET completed = 1 WHERE WOPK = " & WOPK
			Call db.RunSQL(sql,"")
			Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
		Else

			Set rs = db.RunSPReturnRS("MC_GetWorkOrderAssignedLaborForClose",Array(Array("@WOPK", adInteger, adParamInput, 4, WOPK)),"")

			If db.dok Then
				sql = ""
				Do Until rs.eof
					If Not Request.Form("CA_" & rs("pk")) = "" Then
						sql = sql & "Update WOAssignStatus SET completed = 1 WHERE PK = " & rs("pk") & " "
					Else
						sql = sql & "Update WOAssignStatus SET completed = 0 WHERE PK = " & rs("pk") & " "
					End If
					rs.movenext()
				Loop
				If Not sql = "" Then
					'Response.Write sql
					'Response.End
					Call db.RunSQL(sql,"")
					Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If
			End If

		End If

		If Not db.RunSP("MC_CloseWorkOrder", Array(_
			Array("@WOPK", adInteger, adParamInput, 4, WOPK),_
			Array("@requesteddate", adVarChar, adParamInput, 17, SQLdatetimeAT(requesteddate & " " & requestedtime)),_
			Array("@requestedinitials", adChar, adParamInput, 5, requestedinitials),_
			Array("@issueddate", adVarChar, adParamInput, 17, SQLdatetimeAT(issueddate & " " & issuedtime)),_
			Array("@issuedinitials", adChar, adParamInput, 5, issuedinitials),_
			Array("@responded", adBoolean, adParamInput, 1, responded),_
			Array("@respondeddate", adVarChar, adParamInput, 17, SQLdatetimeAT(respondeddate & " " & respondedtime)),_
			Array("@respondedinitials", adChar, adParamInput, 5, respondedinitials),_
			Array("@completed", adBoolean, adParamInput, 1, completed),_
			Array("@completeddate", adVarChar, adParamInput, 17, SQLdatetimeAT(completeddate & " " & completedtime)),_
			Array("@completedinitials", adChar, adParamInput, 5, completedinitials),_
			Array("@completeassignments", adBoolean, adParamInput, 1, completeassignments),_
			Array("@finalized", adBoolean, adParamInput, 1, finalized),_
			Array("@finalizeddate", adVarChar, adParamInput, 17, SQLdatetimeAT(finalizeddate & " " & finalizedtime)),_
			Array("@finalizedinitials", adChar, adParamInput, 5, finalizedinitials),_
			Array("@closed", adBoolean, adParamInput, 1, closed),_
			Array("@closeddate", adVarChar, adParamInput, 17, SQLdatetimeAT(closeddate & " " & closedtime)),_
			Array("@closedinitials", adChar, adParamInput, 5, closedinitials),_
			Array("@laborreport",  adVarchar, adParamInput, 6000, laborreport),_
			Array("@lri", adSmallInt, adParamInput, 2, lri),_
			Array("@txtaccountpk", adInteger, adParamInput, 4, txtaccountpk),_
			Array("@txtaccount",  adVarchar, adParamInput, 25, Trim(Mid(txtaccount,1,25))),_
			Array("@txtaccountdesc",  adVarchar, adParamInput, 50, Trim(Mid(txtaccountdesc,1,50))),_
			Array("@txtAccountAll", adBoolean, adParamInput, 1, txtAccountAll),_
			Array("@txtchargeable", adBoolean, adParamInput, 1, txtchargeable),_
			Array("@txtcategorypk", adInteger, adParamInput, 4, txtcategorypk),_
			Array("@txtcategory",  adVarchar, adParamInput, 25, Trim(Mid(txtcategory,1,25))),_
			Array("@txtcategorydesc",  adVarchar, adParamInput, 50, Trim(Mid(txtcategorydesc,1,50))),_
			Array("@txtCategoryAll", adBoolean, adParamInput, 1, txtCategoryAll),_
			Array("@txtTasks", adBoolean, adParamInput, 1, txtTasks),_
			Array("@txtTaskInitials",  adVarchar, adParamInput, 5, Trim(Mid(txtTaskInitials,1,5))),_
			Array("@txtLabor1", adBoolean, adParamInput, 1, txtLabor1),_
			Array("@txtLabor3", adBoolean, adParamInput, 1, txtLabor3),_
			Array("@txtMaterials", adBoolean, adParamInput, 1, txtMaterials),_
			Array("@txtOtherCost", adBoolean, adParamInput, 1, txtOtherCost),_
			Array("@txtproblempk", adInteger, adParamInput, 4, txtproblempk),_
			Array("@txtproblem",  adVarchar, adParamInput, 25, Trim(Mid(txtproblem,1,25))),_
			Array("@txtproblemdesc",  adVarchar, adParamInput, 50, Trim(Mid(txtproblemdesc,1,50))),_
			Array("@txtfailurepk", adInteger, adParamInput, 4, txtfailurepk),_
			Array("@txtfailure",  adVarchar, adParamInput, 25, Trim(Mid(txtfailure,1,25))),_
			Array("@txtfailuredesc",  adVarchar, adParamInput, 50, Trim(Mid(txtfailuredesc,1,50))),_
			Array("@txtsolutionpk", adInteger, adParamInput, 4, txtsolutionpk),_
			Array("@txtsolution",  adVarchar, adParamInput, 25, Trim(Mid(txtsolution,1,25))),_
			Array("@txtsolutiondesc",  adVarchar, adParamInput, 50, Trim(Mid(txtsolutiondesc,1,50))),_
			Array("@txtfailurewo", adBoolean, adParamInput, 1, txtfailurewo),_
			Array("@txtmeter1reading", adInteger, adParamInput, 4, txtmeter1reading),_
			Array("@txtmeter2reading", adInteger, adParamInput, 4, txtmeter2reading),_
			Array("@txtisup", adBoolean, adParamInput, 1, txtisup),_
			Array("@txtDrawingUpdatesNeeded", adBoolean, adParamInput, 1, txtDrawingUpdatesNeeded),_
			Array("@mode",  adVarchar, adParamInput, 15, mode),_
			Array("@woauth", adChar, adParamInput, 1, GetSession("WOAuth")),_
			Array("@RowVersionUserPK", adInteger, adParamInput, 4, GetSession("UserPK")),_
			Array("@RowVersionInitials",  adVarchar, adParamInput, 5, GetSession("UserInitials")),_
			Array("@RowVersionIPAddress",  adVarchar, adParamInput, 25, GetSession("UserIPAddress")),_
			Array("@txtMyLaborHrs", adVarChar, adParamInput, 15, txtMyLaborHrs),_
			Array("@txtRESPONDEDOVERWRITE", adBoolean, adParamInput, 1, txtRESPONDEDOVERWRITE),_
			Array("@txtCOMPLETEDOVERWRITE", adBoolean, adParamInput, 1, txtCOMPLETEDOVERWRITE),_
			Array("@txtFINALIZEDOVERWRITE", adBoolean, adParamInput, 1, txtFINALIZEDOVERWRITE),_
      Array("@txtMyLaborHrsPK", adInteger, adParamInput, 4, txtMyLaborHrsPK),_
			Array("@txtMyLaborOHrs", adVarChar, adParamInput, 15, txtMyLaborOHrs),_
      Array("@txtFollowUpSingleWO", adBoolean, adParamInput, 1, txtFollowUpSingleWO),_
      Array("@txtFollowUpMultiWO", adBoolean, adParamInput, 1, txtFollowUpMultiWO)_
			),"") Then

			Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

		End If

	Else
		' ************************************************************************************************************
		' CLOSE MULTIPLE WORK ORDERS
		' ************************************************************************************************************
		Dim rs,raforaction,rows,row,wocount
    'Response.Write "WOGroupPK: " & WOGroupPK & "<br>"
    'Response.Write "actionwhere: " & actionwhere & "<br>"
    'Response.Write "mode: " & mode
    'Response.End
    'Response.Write "dberror: " & db.derror & "<br>" 
    
	  db.dok = True
		db.derror = ""

		If mode = "WOGROUP" Then
			If Not Trim(Request.Form("txtWOGroupAll")) = "" And Not WOGroupPK = "-1" And Not WOGroupPK = "" Then
				sql = "SELECT WOPK FROM WO WITH (NOLOCK) WHERE (WOGroupPK = " & WOGroupPK & " OR (WOGroupPK IN (SELECT WOGroupPK FROM WO WITH (NOLOCK) WHERE WOGroupPK = " & WOGroupPK & " AND WOGroupPK > 0 AND WOGroupType = 'M'))) AND WO.IsOpen = 1 ORDER BY WOPK"
			Else
				sql = "SELECT WOPK FROM WO WITH (NOLOCK) " & Replace(actionwhere,"WHERE ","WHERE ( ") & " OR (WOGroupPK IN (SELECT WOGroupPK FROM WO WITH (NOLOCK) " & actionwhere & " AND WOGroupPK > 0 AND WOGroupType = 'M'))) AND WO.IsOpen = 1 ORDER BY WOPK"
			End If
      'Response.Write "sql: <textarea rows=4 cols=60>" & sql & "</textarea><br>"
      'Response.End

			Set rs = db.RunSqlReturnRS(sql,"")

      If db.dok then
			  If Not rs.RecordCount > 0 Then
				  Call CloseObj(rs)
				  Exit Sub
			  Else
			    wocount=rs.RecordCount
				  raforaction = rs.getrows()
				  rows=ubound(raforaction,2)
				  Call CloseObj(rs)
			  End If
      Else
        Response.Write "sql: <textarea rows=4 cols=60>" & sql & "</textarea><br>"
        Response.Write "dberror: <textarea rows=4 cols=60>" & db.derror & "</textarea><br>"
        Response.End
      End If

			Call SetScriptTimeoutTo(60)

			For row = 0 To rows
        dim rh, oh,xh
        rh=0
        oh=0
        xh=0

        'RGJ Start
        'Labor Actuals
        If NullCheck(Request.Form("LaborPK")) <> "" Then
          For X = (LBound(lpk)) To (UBound(lpk))
            workdate = SQLDateTime(wd(X))
            If SplitLabor = 1 Then

              rh = reg(X)/wocount
              oh = ovt(X)/wocount
              xh = oth(X)/wocount

              'Response.Write rh & "<br>"
              'response.Write oh & "<br>"
              'Response.Write xh
              'Response.End

              sql = "INSERT INTO WOLabor (WOPK, LaborPK, LaborID, LaborName, RecordType, LaborType, laborTypeDesc, WorkDate, RegularHours, OvertimeHours, OtherHours, AutoCalcCost, CostRegular, CostOvertime, CostOther, ChargeRate, ChargePercentage, Comments, RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionDate) "+_
              "SELECT " & raforaction(0,row) & ", LaborPK, LaborID, LaborName, 2, LaborType, LaborTypeDesc, '"&workdate&"',"&Round(rh,2)&","&Round(oh,2)&","&Round(xh,2)&",1, CostRegular, CostOvertime, CostOther, ChargeRate, ChargePercentage, '"&SQLEncode(lcom(X))&"', '"&rowuserip&"', "&rowuser&", '"&rowinitials&"', GETDATE() FROM Labor WHERE LaborPK = "&lpk(X)
              'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
              'Response.End
            Else
              sql = "INSERT INTO WOLabor (WOPK, LaborPK, LaborID, LaborName, RecordType, LaborType, laborTypeDesc, WorkDate, RegularHours, OvertimeHours, OtherHours, AutoCalcCost, CostRegular, CostOvertime, CostOther, ChargeRate, ChargePercentage, Comments, RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionDate) "+_
              "SELECT " & raforaction(0,row) & ", LaborPK, LaborID, LaborName, 2, LaborType, LaborTypeDesc, '"&workdate&"',"&Round(reg(X),2)&","&Round(ovt(X),2)&","&Round(oth(X),2)&",1, CostRegular, CostOvertime, CostOther, ChargeRate, ChargePercentage, '"&SQLEncode(lcom(X))&"', '"&rowuserip&"', "&rowuser&", '"&rowinitials&"', GETDATE() FROM Labor WHERE LaborPK = "&lpk(X)
              'Response.Write "<textarea rows=6 cols=100>"&sql&"</textarea>"
              'Response.End
            End If
            'Response.Write "labor insert"
            'Response.End

            Call db.RunSQL(sql,"")
            Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
          Next

          'Update the WO cost fields
          Call db.RunSP("MC_CalcWorkOrder",Array(Array("@WOPK", adInteger, adParamInput, 4, WOPK)),"")
        End If

        'RGJ END


				If completeassignments = 1 Then
					sql = "Update WOAssignStatus SET completed = 1 WHERE WOPK = " & raforaction(0,row)
          'Response.Write "Assignments update"
          'Response.End					
          Call db.RunSQL(sql,"")
					Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				If Not db.RunSP("MC_CloseWorkOrder", Array(_
					Array("@WOPK", adInteger, adParamInput, 4, raforaction(0,row)),_
					Array("@requesteddate", adVarChar, adParamInput, 17, SQLdatetimeAT(requesteddate & " " & requestedtime)),_
					Array("@requestedinitials", adChar, adParamInput, 5, requestedinitials),_
					Array("@issueddate", adVarChar, adParamInput, 17, SQLdatetimeAT(issueddate & " " & issuedtime)),_
					Array("@issuedinitials", adChar, adParamInput, 5, issuedinitials),_
					Array("@responded", adBoolean, adParamInput, 1, responded),_
					Array("@respondeddate", adVarChar, adParamInput, 17, SQLdatetimeAT(respondeddate & " " & respondedtime)),_
					Array("@respondedinitials", adChar, adParamInput, 5, respondedinitials),_
					Array("@completed", adBoolean, adParamInput, 1, completed),_
					Array("@completeddate", adVarChar, adParamInput, 17, SQLdatetimeAT(completeddate & " " & completedtime)),_
					Array("@completedinitials", adChar, adParamInput, 5, completedinitials),_
					Array("@completeassignments", adBoolean, adParamInput, 1, completeassignments),_
					Array("@finalized", adBoolean, adParamInput, 1, finalized),_
					Array("@finalizeddate", adVarChar, adParamInput, 17, SQLdatetimeAT(finalizeddate & " " & finalizedtime)),_
					Array("@finalizedinitials", adChar, adParamInput, 5, finalizedinitials),_
					Array("@closed", adBoolean, adParamInput, 1, closed),_
					Array("@closeddate", adVarChar, adParamInput, 17, SQLdatetimeAT(closeddate & " " & closedtime)),_
					Array("@closedinitials", adChar, adParamInput, 5, closedinitials),_
					Array("@laborreport",  adVarchar, adParamInput, 6000, Trim(Mid(laborreport,1,6000))& " "),_
					Array("@lri", adSmallInt, adParamInput, 2, lri),_
					Array("@txtaccountpk", adInteger, adParamInput, 4, txtaccountpk),_
					Array("@txtaccount",  adVarchar, adParamInput, 25, Trim(Mid(txtaccount,1,25))),_
					Array("@txtaccountdesc",  adVarchar, adParamInput, 50, Trim(Mid(txtaccountdesc,1,50))),_
					Array("@txtAccountAll", adBoolean, adParamInput, 1, txtAccountAll),_
					Array("@txtchargeable", adBoolean, adParamInput, 1, txtchargeable),_
					Array("@txtcategorypk", adInteger, adParamInput, 4, txtcategorypk),_
					Array("@txtcategory",  adVarchar, adParamInput, 25, Trim(Mid(txtcategory,1,25))),_
					Array("@txtcategorydesc",  adVarchar, adParamInput, 50, Trim(Mid(txtcategorydesc,1,50))),_
					Array("@txtCategoryAll", adBoolean, adParamInput, 1, txtCategoryAll),_
					Array("@txtTasks", adBoolean, adParamInput, 1, txtTasks),_
					Array("@txtTaskInitials",  adVarchar, adParamInput, 5, Trim(Mid(txtTaskInitials,1,5))),_
					Array("@txtLabor1", adBoolean, adParamInput, 1, txtLabor1),_
					Array("@txtLabor3", adBoolean, adParamInput, 1, txtLabor3),_
					Array("@txtMaterials", adBoolean, adParamInput, 1, txtMaterials),_
					Array("@txtOtherCost", adBoolean, adParamInput, 1, txtOtherCost),_
					Array("@txtproblempk", adInteger, adParamInput, 4, txtproblempk),_
					Array("@txtproblem",  adVarchar, adParamInput, 25, Trim(Mid(txtproblem,1,25))),_
					Array("@txtproblemdesc",  adVarchar, adParamInput, 50, Trim(Mid(txtproblemdesc,1,50))),_
					Array("@txtfailurepk", adInteger, adParamInput, 4, txtfailurepk),_
					Array("@txtfailure",  adVarchar, adParamInput, 25, Trim(Mid(txtfailure,1,25))),_
					Array("@txtfailuredesc",  adVarchar, adParamInput, 50, Trim(Mid(txtfailuredesc,1,50))),_
					Array("@txtsolutionpk", adInteger, adParamInput, 4, txtsolutionpk),_
					Array("@txtsolution",  adVarchar, adParamInput, 25, Trim(Mid(txtsolution,1,25))),_
					Array("@txtsolutiondesc",  adVarchar, adParamInput, 50, Trim(Mid(txtsolutiondesc,1,50))),_
					Array("@txtfailurewo", adBoolean, adParamInput, 1, txtfailurewo),_
					Array("@txtmeter1reading", adInteger, adParamInput, 4, txtmeter1reading),_
					Array("@txtmeter2reading", adInteger, adParamInput, 4, txtmeter2reading),_
					Array("@txtisup", adBoolean, adParamInput, 1, txtisup),_
					Array("@txtDrawingUpdatesNeeded", adBoolean, adParamInput, 1, txtDrawingUpdatesNeeded),_
					Array("@mode",  adVarchar, adParamInput, 15, mode),_
					Array("@woauth", adChar, adParamInput, 1, GetSession("WOAuth")),_
					Array("@RowVersionUserPK", adInteger, adParamInput, 4, GetSession("UserPK")),_
					Array("@RowVersionInitials",  adVarchar, adParamInput, 5, GetSession("UserInitials")),_
					Array("@RowVersionIPAddress",  adVarchar, adParamInput, 25, GetSession("UserIPAddress")),_
    			Array("@txtMyLaborHrs", adVarChar, adParamInput, 15, txtMyLaborHrs),_
	        Array("@txtRESPONDEDOVERWRITE", adBoolean, adParamInput, 1, txtRESPONDEDOVERWRITE),_
	        Array("@txtCOMPLETEDOVERWRITE", adBoolean, adParamInput, 1, txtCOMPLETEDOVERWRITE),_
	        Array("@txtFINALIZEDOVERWRITE", adBoolean, adParamInput, 1, txtFINALIZEDOVERWRITE),_
          Array("@txtMyLaborHrsPK", adInteger, adParamInput, 4, txtMyLaborHrsPK),_
	        Array("@txtMyLaborOHrs", adVarChar, adParamInput, 15, txtMyLaborOHrs),_
          Array("@txtFollowUpSingleWO", adBoolean, adParamInput, 1, txtFollowUpSingleWO),_
          Array("@txtFollowUpMultiWO", adBoolean, adParamInput, 1, txtFollowUpMultiWO)_
         ),"") Then
            Response.Write "WO Close"
            Response.End
					Call dok_check_popup(db,"Work Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					Exit For

				End If
			Next

			ResetScriptTimeOut

		End If

	End If

	'sql = ""

	'Call db.runSQL(sql,"")

	If Not db.dok Then
		errormessage = db.derror
	End If

	' ===============================================================================
	'Call aspdebug()
	%>
	<script language="javascript">
		//alert("<% =sql %>");
		//alert('<% =JSEncode(errormessage) %>');
	</script>
	<% =errormessage %>
	</body>
	</html>
	<%
	Response.End

End Sub

'JReed Code for FollowupWO display
Sub WO_CLOSE_SHOWONDEMAND_FOLLOWUP
%>
<!-- Insert table to create follow-up WO -->	
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-top:15;">
    <tr>
        <td class="fieldsheader" style="border-bottom:0 solid #CCCCCC;">
            <img SRC="../../images/icons/shareitxp_g.gif" style="margin-right:5;" width="15" height="13">Follow-up Work
        </td>
        <td align="right" class="fieldsheader" style="border-bottom:0 solid #CCCCCC;">
            <img class="mcbutton" id="followup" border="0" src="../../images/button_new.gif" onclick="openOnDemandFollowupWO(this);event.cancelBubble=true;" WIDTH="80" HEIGHT="15">									    
            <img style="display:none;" class="mcbutton" border="0" src="../../images/button_action.gif" onclick="top.showMenu('wofailmenu',top.recordkey,true,top.fraTopic,'FORCELEFT');this.blur();return false" WIDTH="80" HEIGHT="15">									    
        </td>
    </tr>																		
    </table> 																	
    <!-- Follow-up Work Table -->
    <table id="ofw1" modid="WO" style="margin-top:0;" border="1" cellspacing="0" cellpadding="1" width="100%" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC">
        <tbody id="ofw1body" mcindex="0">			
        <tr style="display:none;">
            <td style="display:none;" class="tableheadercell" valign="top" width="10" height="20">
                <img alt="Click to select all" hspace="4" style="cursor:hand;" src="../../images/checkbox.gif" border="0" width="13" height="11" onclick="top.select_all(top.fraTopic.document,'fw1')">
            </td> 
            <td style="padding-left:5px;" width="19" class="tableheadercellleft" valign="top">
            <img src="../../images/icons/wo2_g.gif" border="0" width="12" height="13">
            </td>
            <td class="tableheadercellleft" valign="top">Work Order</td>
            <td class="tableheadercellleft" valign="top">Target Date</td>
        </tr>
        <tr STYLE="display: none">		
            <td style="display:none;" class="tablecboxcell" valign="top">
                <img class="lookupicon" src="../../images/undo2.gif" onclick="top.mcTRUndo(top.findRow(this));" WIDTH="14" HEIGHT="16">
            </td>
            <td style="display:none;"></td>
            <td class="tabledatacellleft" valign="top" nowrap>
            </td>
            <td class="tabledatacellleft" valign="top" nowrap>
            </td>
        </tr>
        <tr viewtemplate="Y" oncontextmenu="top.mcTR_OnContextMenu(this,true,false);return false;" STYLE="display: none" onclick="top.doviewedit('WO',this.mccontextkey,top.fraTopic);" onmouseover="ofwTR_OnMouseOver(this);" onmouseout="ofwTR_OnMouseOut(this);">
            <td style="display:none;" class="tablecboxcell" valign="top">
                <input type="checkbox" name="checkedDataList_cc1ROW_ID" class="mccheckbox" onclick="top.mctr_checked(this);">
            </td> 
            <td width="19" class="tabledatacellleft" valign="top">
            </td>
            <td class="tabledatacellleft" valign="top">
            </td>
            <td class="tabledatacellleft" valign="top">
            </td>
        </tr>				
    	
        </tbody>
    </table>
    <!-- Follow-up Work Table End -->
<%
End Sub

'RGJ BEGIN - Start building Subs that are content items
'Show Status Dates
Sub WO_CLOSE_SHOWSTATUSDATES
    %>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family:Arial; font-size:9pt; color:#000000;" width="100%">
    <tr>
      <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;" colspan="2">
        <img SRC="../../images/icons/calendarxpl_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="15">Status Dates
      </td>
    </tr>
	  <tr>
		  <td height="5" colspan="2"></td>
	  </tr>
    <tr id="requestedbox" style="display:<% If Not mode = "WO" Then %>none<% End If %>;">
      <td valign="top" style="padding-left:2px; padding-top:2px;">
        <img style="margin-right:3px;" src="../../images/icons/status_phoned_g.gif" border="0" >Requested:&nbsp;&nbsp;
      </td>
      <td>
        <table cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td valign="top" style="display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="D" mcName="Requested Date" mcRequired="Y" type="text" name="txtRequestedDate" id="txtRequestedDate" value="<% =Requesteddate %>" size="10" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgRequestedDate" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calendar','Calendar',172,160,this,txtRequestedDate)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtRequestedDateErr" class="mc_lookupdesc"></span>
            </td>
            <td valign="top" style="padding-left:10px; display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="T" mcName="Requested Time" mcRequired="Y" type="text" name="txtRequestedTime" id="txtRequestedTime" value="<% =Requestedtime %>" size="9" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgRequestedTime" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('timepopup','Select Time',267,205,this,txtRequestedTime)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtRequestedTimeErr" class="mc_lookupdesc"></span>
            </td>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;">
  	          <span style="padding-right:4px;">by:</span><input type="text" name="txtRequestedInitials" id="txtRequestedInitials" value="<% =RequestedInitials %>" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>class="disabled" readonly<%Else%>class="normal" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"<%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="5">
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr id="issuedbox" style="display:<% If Not mode = "WO" Then %>none<% End If %>;">
      <td valign="top" style="padding-left:2px; padding-top:2px;">
        <img style="margin-right:3px;" src="../../images/icons/status_inprogress_g.gif" border="0">Issued:&nbsp;&nbsp;
      </td>
      <td>
        <table cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td valign="top" style="display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="D" mcName="Issued Date" mcRequired="Y" type="text" name="txtIssuedDate" id="txtIssuedDate" value="<% =Issueddate %>" size="10" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgIssuedDate" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calendar','Calendar',172,160,this,txtIssuedDate)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtIssuedDateErr" class="mc_lookupdesc"></span>
            </td>
            <td valign="top" style="padding-left:10px; display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="T" mcName="Issued Time" mcRequired="Y" type="text" name="txtIssuedTime" id="txtIssuedTime" value="<% =Issuedtime %>" size="9" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgIssuedTime" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('timepopup','Select Time',267,205,this,txtIssuedTime)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtIssuedTimeErr" class="mc_lookupdesc"></span>
            </td>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;">
  	          <span style="padding-right:4px;">by:</span><input type="text" name="txtIssuedInitials" id="txtIssuedInitials" value="<% =IssuedInitials %>" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>class="disabled" readonly<%Else%>class="normal" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"<%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="5">
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr id="respondedbox" style="display:<% If Not AccessToRespond Then %>none<% End If %>">
      <td valign="top" style="padding-top:0px;">
        <input type="hidden" value="<% If responded Then %>Y<% Else %>N<% End If %>" name="txtResponded">
        <img onclick="processaction(this);" style="margin-right:8px; cursor:hand;" name="respondedimg" id="respondedimg" src="../../images/button_respond_<% If responded Then %>on<% Else %>off<% End If %>.gif">
      </td>
      <td>
        <table id="respondeddata" style="display:<% If Not responded Then %>none<% End If %>;" cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td valign="top" style="display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="D" mcName="Responded Date" mcRequired="Y" type="text" name="txtRespondedDate" id="txtRespondedDate" value="<% =respondeddate %>" size="10" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgRespondedDate" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calendar','Calendar',172,160,this,txtRespondedDate)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtRespondedDateErr" class="mc_lookupdesc"></span>
            </td>
            <td valign="top" style="padding-left:10px; display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="T" mcName="Responded Time" mcRequired="Y" type="text" name="txtRespondedTime" id="txtRespondedTime" value="<% =respondedtime %>" size="9" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgRespondedTime" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('timepopup','Select Time',267,205,this,txtRespondedTime)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtRespondedTimeErr" class="mc_lookupdesc"></span>
            </td>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;">
  	          <span style="padding-right:4px;">by:</span><input type="text" name="txtRespondedInitials" id="txtRespondedInitials" value="<% =RespondedInitials %>" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>class="disabled" readonly<%Else%>class="normal" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"<%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="5">
            </td>
            <%If Not mode = "WO" Then %>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;" onclick="txtRESPONDEDOVERWRITE.click();">
	            <input style="margin-right:1px;" onclick="this.blur();" type="checkbox" value="ON" name="txtRESPONDEDOVERWRITE" tabindex="">Overwrite Existing Values?
            </td>
            <%End If %>
          </tr>
        </table>
      </td>
    </tr>
    <tr id="completedbox" style="display:<% If Not AccessToComplete Then %>none<% End If %>">
      <td valign="top" style="padding-top:0px;">
        <input type="hidden" value="<% If completed Then %>Y<% Else %>N<% End If %>" name="txtCompleted">
        <img onclick="processit(this);" style="margin-right:8px; cursor:hand;" name="completedimg" id="completedimg" src="../../images/button_complete_<% If completed Then %>on<% Else %>off<% End If %>.gif">
      </td>
      <td>
        <table id="completeddata" style="display:<% If Not completed Then %>none<% End If %>;" cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td valign="top" style="display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="D" mcName="Completed Date" mcRequired="Y" type="text" name="txtCompletedDate" id="txtCompletedDate" value="<% =Completeddate %>" size="10" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgCompletedDate" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calendar','Calendar',172,160,this,txtCompletedDate)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtCompletedDateErr" class="mc_lookupdesc"></span>
            </td>
            <td valign="top" style="padding-left:10px; display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="T" mcName="Completed Time" mcRequired="Y" type="text" name="txtCompletedTime" id="txtCompletedTime" value="<% =Completedtime %>" size="9" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgCompletedTime" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('timepopup','Select Time',267,205,this,txtCompletedTime)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtCompletedTimeErr" class="mc_lookupdesc"></span>
            </td>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;">
  	          <span style="padding-right:4px;">by:</span><input type="text" name="txtCompletedInitials" id="txtCompletedInitials" value="<% =CompletedInitials %>" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>class="disabled" readonly<%Else%>class="normal" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"<%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="5">
            </td>
            <%If Not mode = "WO" Then %>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;" onclick="txtCOMPLETEDOVERWRITE.click();">
	            <input style="margin-right:1px;" onclick="this.blur();" type="checkbox" value="ON" name="txtCOMPLETEDOVERWRITE" tabindex="">Overwrite Existing Values?
            </td>
            <%End If %>
          </tr>
        </table>
      </td>
    </tr>
    <tr id="finalizedbox" style="display:<% If Not AccessToFinalize Then %>none<% End If %>">
      <td valign="top" style="padding-top:0px;">
        <input type="hidden" value="<% If finalized Then %>Y<% Else %>N<% End If %>" name="txtFinalized">
        <img onclick="processit(this);" style="margin-right:8px; cursor:hand;" name="finalizedimg" id="finalizedimg" src="../../images/button_finalize_<% If finalized Then %>on<% Else %>off<% End If %>.gif">
      </td>
      <td>
        <table id="finalizeddata" style="display:<% If Not finalized Then %>none<% End If %>;" cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td valign="top" style="display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="D" mcName="Finalized Date" mcRequired="Y" type="text" name="txtFinalizedDate" id="txtFinalizedDate" value="<% =Finalizeddate %>" size="10" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgFinalizedDate" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calendar','Calendar',172,160,this,txtFinalizedDate)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtFinalizedDateErr" class="mc_lookupdesc"></span>
            </td>
            <td valign="top" style="padding-left:10px;display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="T" mcName="Finalized Time" mcRequired="Y" type="text" name="txtFinalizedTime" id="txtFinalizedTime" value="<% =Finalizedtime %>" size="9" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgFinalizedTime" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('timepopup','Select Time',267,205,this,txtFinalizedTime)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtFinalizedTimeErr" class="mc_lookupdesc"></span>
            </td>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;">
  	          <span style="padding-right:4px;">by:</span><input type="text" name="txtFinalizedInitials" id="txtFinalizedInitials" value="<% =FinalizedInitials %>" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>class="disabled" readonly<%Else%>class="normal" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"<%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="5">
            </td>
            <%If Not mode = "WO" Then %>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;" onclick="txtFINALIZEDOVERWRITE.click();">
	            <input style="margin-right:1px;" onclick="this.blur();" type="checkbox" value="ON" name="txtFINALIZEDOVERWRITE" tabindex="">Overwrite Existing Values?
            </td>
            <%End If %>
          </tr>
        </table>
      </td>
    </tr>
    <tr id="closedbox" style="display:<% If Not AccessToClose Then %>none<% End If %>">
      <td valign="top" style="padding-top:0px;">
        <input type="hidden" value="<% If closed Then %>Y<% Else %>N<% End If %>" name="txtClosed">
        <img onclick="processit(this);" style="margin-right:8px; cursor:hand;" name="closedimg" id="closedimg" src="../../images/button_close_<% If closed Then %>on<% Else %>off<% End If %>.gif">
      </td>
      <td>
        <table id="closeddata" style="display:<% If Not closed Then %>none<% End If %>;" cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td valign="top" style="display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="D" mcName="Closed Date" mcRequired="Y" type="text" name="txtClosedDate" id="txtClosedDate" value="<% =Closeddate %>" size="10" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgClosedDate" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calendar','Calendar',172,160,this,txtClosedDate)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtClosedDateErr" class="mc_lookupdesc"></span>
            </td>
            <td valign="top" style="padding-left:10px; display:<% If Not AccessToStatusDateTime Then %>none<% End If %>;">
              <input mcType="T" mcName="Closed Time" mcRequired="Y" type="text" name="txtClosedTime" id="txtClosedTime" value="<% =Closedtime %>" size="9" maxlength="12" TABINDEX="" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>CLASS="disabled" readonly<%Else%>CLASS="required" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"<%End If%>><%If WO_CLOSE_STATUSDATESREADONLY = "No" Then%><img id="imgClosedTime" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('timepopup','Select Time',267,205,this,txtClosedTime)" class="lookupicon" WIDTH="16" HEIGHT="20"><%End If%>
              <span id="txtClosedTimeErr" class="mc_lookupdesc"></span>
            </td>
            <td style="cursor:hand; padding-top:0px; padding-left:10px; font-family:Arial; font-size:9pt; color:#000000;">
  	          <span style="padding-right:4px;">by:</span><input type="text" name="txtClosedInitials" id="txtClosedInitials" value="<% =ClosedInitials %>" <%If WO_CLOSE_STATUSDATESREADONLY = "Yes" Then%>class="disabled" readonly<%Else%>class="normal" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"<%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="5">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<%
End Sub

'Show Labor Report
Sub WO_CLOSE_SHOWLABORREPORT
%>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-family:Arial; font-size:9pt; color:#000000">
		<tr>
			<td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
		    <img SRC="../../images/icons/paperxp_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="15">Labor Report
			</td>
			<td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;" align="right">
				
        <table cellspacing="0" cellpadding="0" border="0">
					<tr>
            <!--            
						<td style="cursor:hand;" onclick="addNewNote();" onmouseover="$(this).css('background-color','#FFFF99');" onmouseout="$(this).css('background-color','transparent');">
              <table>
                <tr>
                  <td align="right"><img border="0" src="../../images/icons/addnote.gif"  width="17" height="17" /></td>
                  <td valign="middle" style="padding-right:15px; font-family: Arial; font-size:8pt; font-weight:bold;">&nbsp;Add Note</td>
                </tr>
              </table>
						</td>
            -->
            <td>
            <%'If WO_CLOSE_LABORREPORTOPTIONS = "3" Or WO_CLOSE_LABORREPORTOPTIONS = "4" Then %>
              <!--<img class="mcbutton" border="0" src="../../images/button_add.gif" onclick="addNewNote('1');" WIDTH="80" HEIGHT="15">-->
            <%'ElseIf WO_CLOSE_LABORREPORTOPTIONS = "2" Then%>
              <!--<img class="mcbutton" border="0" src="../../images/button_add.gif" onclick="addNewNote('2');" WIDTH="80" HEIGHT="15">-->
            <%'Else%>
              <img class="mcbutton" border="0" src="../../images/button_addarrow.gif" onclick="top.showpopup('actions','Actions',266,100,this,txtreport);" WIDTH="80" HEIGHT="15">
            <%'End If%>
            </td>
					</tr>
		    </table>
			</td>
		</tr>
		<tr>
	  	<td height="2" colspan="2">
			</td>
		</tr>
		<tr>
			<td colspan="2">
				<textarea maxlength="6000" mcType="C" <%If WO_CLOSE_SHOWLABORREPORT_REQ = "Yes" Then%>class="required1" <%Else%>class="normal" <%End If%> mcName="txtreport" id="txtreport" name="txtreport" wrap="hard" style="margin-top:5; width: 100%; height: 100;" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this);" onkeydowns="checkForDT();" onblur="top.fieldblur(this);validateTextarea(this);" onChange="top.fieldvalid(this)" rows="1" cols="20" tabindex=""><% =laborreport %></textarea>
				<div id="option_lri" style="margin-top:3px;<% If mode = "WO" Then %>display:none;<% End If %>">
					Existing Labor Reports: <input style="position:relative; top:-1px; height:13px;" onclick="this.blur();" type="radio" checked name="lri" value="1"> <span style="cursor:hand;" onclick="document.mcform.lri[0].click();">Do Nothing</span> <input style="position:relative; top:-1px; height:13px;" onclick="this.blur();" type="radio" name="lri" value="2"> <span style="cursor:hand;" onclick="document.mcform.lri[1].click();">Append</span> <input style="position:relative; top:-1px; height:13px;" type="radio" onclick="this.blur();" name="lri" value="3"> <span style="cursor:hand;" onclick="document.mcform.lri[2].click();">Overwrite</span>
				</div>
			</td>
		</tr>
	</table>
<%End Sub

'Show Actions
Sub WO_CLOSE_SHOWACTIONS
%>
  <table id="oola2header" style="font-family:Arial; font-size:9pt; color:#000000" border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
        <img SRC="../../images/icons/status_generated_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="15">Actions
      </td>
    </tr>
	  <tr>
		  <td height="5"></td>
	  </tr>
    <tr>
	    <td>
			  <table id="ota1header" style="margin-top:0;font-family:Arial; font-size:9pt; color:#000000" border="0" cellpadding="0" cellspacing="0" width="90%">
				  <tr style="display:none;">
					  <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
				      <img SRC="../../images/icons/tasks_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="12">Task Check List
					  </td>
				  </tr>
          <tr style="display:;">
					  <td height="5"></td>
    		  </tr>
          <tr>
					  <td>
						  <span style="cursor:hand;" onclick="txtTasks.click();">
						  <img SRC="../../images/icons/tasks_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="12"><input onclick="this.blur();top.enabledisable();" type="checkbox" value="ON" name="txtTasks" tabindex="">
						  Set All Tasks Complete</span><span style="padding-left:25px;padding-right:4px;">Initials:</span><input disabled type="text" name="txtTaskInitials" id="txtTaskInitials" value="<% =TaskDefaultInitials %>" class="normal" mcType="C" maxlength="25" mcRequired="N" tabindex="" size="5" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);">
				    </td>
				  </tr>
			  </table>
      </td>
	  </tr>
	  <tr>
		  <td onclick="txtLabor3.click();" style="cursor:hand;">
        <img SRC="../../images/icons/laborsm_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="13"><input onclick="txtLabor1.checked = false;top.enabledisable();this.blur();" type="checkbox" value="ON" name="txtLabor3" tabindex="">
        Set Actual Labor Hours equal to Estimated Labor Hours
      </td>
	  </tr>
	  <tr>
		  <td onclick="txtLabor1.click();" style="cursor:hand;">
  		  <img SRC="../../images/icons/laborsm_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="13"><input onclick="txtLabor3.checked = false;top.enabledisable();this.blur();" type="checkbox" value="ON" name="txtLabor1" tabindex="">
			  Set Actual Labor Hours equal to Assigned Labor Hours
      </td>
	  </tr>
	  <% If False Then %>
	  <tr>
		  <td onclick="txtLabor2.click();" style="cursor:hand;">
        <img SRC="../../images/icons/laborsm_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="13"><input onclick="txtLabor1.checked = false;top.enabledisable();this.blur();" type="checkbox" value="ON" name="txtLabor2" tabindex="">Set Actual
        Labor to a total of
        <input disabled mcType="N" class="requiredright" type="text" name="txtLaborHrs" size="1" style="width:32" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" value="0" tabindex=""><img id="imgLaborHrs" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtLaborHrs);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
        hour(s) (spread over all Assigned Labor)
      </td>
	  </tr>
	  <% End If %>
  </table>
  <table id="oma2header" style="margin-top:0;font-family:Arial; font-size:9pt; color:#000000" border="0" cellpadding="0" cellspacing="0" width="90%">
    <tr style="display:none;">
		  <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
			  <img SRC="../../images/icons/box3d_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="13">Materials
		  </td>
	  </tr>
	  <tr style="display:none;">
  	  <td height="5"></td>
	  </tr>
	  <tr>
		  <td onclick="txtMaterials.click();" style="cursor:hand;">
			  <img SRC="../../images/icons/box3d_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="13"><input onclick="this.blur();" type="checkbox" value="ON" name="txtMaterials" tabindex=""> Set
        Actual Materials equal to Estimated Materials
      </td>
	  </tr>
  </table>
  <table id="oot2header" style="margin-top:0;font-family:Arial; font-size:9pt; color:#000000" border="0" cellpadding="0" cellspacing="0" width="90%">
	  <tr style="display:none;">
		  <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
			  <img SRC="../../images/icons/worldphotoxp_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="13">Other Costs
		  </td>
	  </tr>
	  <tr style="display:none;">
		  <td height="5"></td>
	  </tr>
	  <tr>
		  <td onclick="txtOtherCost.click();" style="cursor:hand;">
			  <img SRC="../../images/icons/worldphotoxp_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="13"><input onclick="this.blur();" type="checkbox" value="ON" name="txtOtherCost" tabindex=""> Set
        Actual Other Costs equal to Estimated Other Costs
      </td>
	  </tr>
  </table>
  <div id="option_drawings" style="position:relative;top:-2px;<% If assetexists or Not mode = "WO" Then %>display:;<% Else %>display:none;<% End If %> ">
    <table id="drawingstable" border="0" cellpadding="0" cellspacing="0" style="margin-top:0;font-family:Arial; font-size:9pt; color:#000000">
	    <tr style="display:none;">
  	    <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
	        &nbsp;<img SRC="../../images/icons/scenicxp_g.gif" HEIGHT="13" style="margin-right:5px;">Asset Drawings
		    </td>
	    </tr>
      <tr>
		    <td height="5" style="display:none;"></td>
	    </tr>
      <tr>
		    <td style="padding-left:1px; cursor:hand; font-family:Arial; font-size:9pt; color:#000000;" onclick="txtDrawingUpdatesNeeded.click();">
			    <img SRC="../../images/icons/scenicxp_g.gif" HEIGHT="13" style="margin-right:5px;"><input onclick="this.blur();event.cancelBubble = true;" type="checkbox" value="ON" name="txtDrawingUpdatesNeeded" tabindex="">
			    Asset Drawing Updates Needed
		    </td>
      </tr>
    </table>
  </div>
  <div id="option_downtime" style="position:relative; top:-2px; <% If assetexists or Not mode = "WO" Then %>display:;<% Else %>display:none;<% End If %>">
	  <table id="downtimetable" border="0" cellpadding="0" cellspacing="0" style="margin-top:0;font-family:Arial; font-size:9pt; color:#000000">
		  <tr style="display:none;">
			  <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
		      &nbsp;<img SRC="../../images/shutdown_all.gif" WIDTH="15" HEIGHT="13"><% If txtisup and mode = "WO" Then %>Downtime<% Else %>Asset Status<% End If %>
			  </td>
		  </tr>
      <tr style="display:none;">
			  <td height="5"></td>
		  </tr>
      <tr>
			  <td onclick="txtDownTime.click();" style="cursor:hand;">
          &nbsp;<img SRC="../../images/shutdown_all.gif" WIDTH="15" HEIGHT="13">&nbsp;<input onclick="top.enabledisable();this.blur();" type="checkbox" value="ON" name="txtDownTime" tabindex="">
          <span id="option_downtime2">
          <% If txtIsUp and mode = "WO" Then %>
            Set Downtime
          <% Else %>
            Return Asset to Service <% If Not mode = "WO" Then %> (if Shutdown)<% End If %>
          <% End If %>
          </span>
        </td>
			</tr>
	  </table>
	</div>
<%End Sub

'Show Assignments
Sub WO_CLOSE_SHOWASSIGNMENTS
%>
  <%
  If getassignments Then
    Set rs = db.RunSPReturnRS("MC_GetWorkOrderAssignedLaborForClose",Array(Array("@WOPK", adInteger, adParamInput, 4, WOPK)),"")
    If db.dok Then
	    If Not rs.eof Then
  %><table cellspacing="0" cellpadding="0" border="0" width="100%" id="laborAssigments">
	  <tr>
		  <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;" colspan="2">
	      <img SRC="../../images/icons/status_requested_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="15">Assignments
		  </td>
	  </tr>
    <tr>
		  <td height="5"></td>
	  </tr>
    <tr>
	    <td id="CompleteAssignmentsBox" onclick="document.mcform.CompleteAssignments.click();" style="cursor:hand; border-bottom: 1 solid #A2A2A2; padding-top:0px;">
		    <input onclick="complete_assignments();" style="width:13px;" type="checkbox" name="CompleteAssignments" value="ON"> <span id="CompleteAssignmentsText" style="font-weight:bold; font-family:arial; font-size:8pt; color:royalblue;">Completed Assignments:</span>
		  </td>
	  </tr>
	  <tr id="WOAssignmentsBox">
		  <td align="right" style="padding-top:5px;">
		    <% If rs.recordcount > 3 Then %>
			  <div style="height:70px; width:400px; overflow-x:auto;">
			  <% Else %>
			  <div style="height:70px; width:100%; overflow-x:auto;">
			  <% End If %>
			    <table cellspacing="0" cellpadding="0" border="0">
				    <tr>
				    <%
				    Dim GotTaskDefaultInitials, acnt
				    acnt=1
				    GotTaskDefaultInitials = False
				    Do Until rs.eof
				      If rs("IsAssigned") and Not GotTaskDefaultInitials Then
						    TaskDefaultInitials = rs("Initials")
						    GotTaskDefaultInitials = True
					    End If
				    %>
					    <td onclick="document.mcform.CA_<% =rs("PK") %>.click();" nowrap valign="top" align="center" style="cursor:hand; font-family:arial; font-size:8pt; color:gray; padding-right:15px;">
					      <% If NullCheck(rs("photo")) = "" Then %>
						    <input class="laborassign" onclick="if (this.checked == false) {document.mcform.CompleteAssignments.checked = false;};checkAssignments();" oldcheckedvalue="<% =LCase(rs("completed")) %>" type="checkbox" <% If rs("completed") Then %>checked <% End If %>name="CA_<% =rs("PK") %>" value="ON"><img style="border:solid 2 #D5D5D5; margin-right:10px;" src="../../images/labor_nophotoxp3_tab.jpg" border="0"><br><% =NullCheck(rs("LaborName")) %>
						    <% Else %>
						    <input class="laborassign" onclick="if (this.checked == false) {document.mcform.CompleteAssignments.checked = false;};checkAssignments();" oldcheckedvalue="<% =LCase(rs("completed")) %>" type="checkbox" <% If rs("completed") Then %>checked <% End If %>name="CA_<% =rs("PK") %>" value="ON"><img style="border:solid 2 #B5B5B5; margin-right:10px;" src="<% =Application("ImageServer") & Replace(NullCheck(rs("Photo")),"_main.","_tab.") %>" border="0"><br><% =NullCheck(rs("LaborName")) %>
						    <% End If %>
					    </td>
				    <%
				      acnt = acnt + 1
					    rs.MoveNext()
				    Loop
				    rs.close()
				    SET rs = nothing
				    %>
				    </tr>
			    </table>
		    </div>
	    </td>
    </tr>
  </table>
  <%
      Else
	      getassignments = False
      End If
    Else
		  getassignments = False
		  errormessage = "There was a problem accessing the Work Order record. Please contact your maintenance manager for support.<br><br>" & db.derror
	  End If
  End If
  If Not getassignments Then
  %>
  <table cellspacing="0" cellpadding="0" border="0">
	  <tr>
	    <td id="CompleteAssignmentsBox" onclick="document.mcform.CompleteAssignments.click();" style="cursor:hand; border-bottom: 0 solid #A2A2A2; padding-top:15px;">
			  <input style="width:13px;" type="checkbox" name="CompleteAssignments" value="ON"> <span id="CompleteAssignmentsText" style="font-weight:bold; font-family:arial; font-size:8pt; color:royalblue;">Set All Assignments Completed</span>
		  </td>
	  </tr>
	  <tr id="WOAssignmentsBox">
		  <td></td>
	  </tr>
  </table>
  <% End If %>
<%End Sub

'Show Failure Analysis
Sub WO_CLOSE_SHOWFAILUREANALYSIS
%>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-family:Arial; font-size:9pt; color:#000000">
	  <tr>
		  <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;" colspan="2">
	      <img SRC="../../images/icons/bellsystem_g.gif" style="margin-right:5;" WIDTH="14" HEIGHT="15">Failure Analysis
		  </td>
	  </tr>
    <tr>
		  <td colspan="2" height="5"></td>
    </tr>
    <tr>
  	  <td valign="top">
			  <table border="0" cellpadding="0" cellspacing="0" style="font-family:Arial; font-size:9pt; color:#000000">
				  <tr>
				    <td nowrap valign="top" style="padding-top:3;">Problem: &nbsp;</td>
				    <td valign="top">
						  <input type="text" name="txtProblem" id="txtProblem" value="<% =txtProblem %>" <%If WO_CLOSE_SHOWFAILUREANALYSIS_PREQ = "Yes" Then%>class="required1" <%Else%>class="normal" <%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="14" onChange="top.dovalid('FA',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" fkm="FA"><img id="imgProblem" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('FA_P',txtProblem,'WO')" class="lookupicon" WIDTH="16" HEIGHT="20">
						  <span id="txtProblemDesc" class="mc_lookupdesc"><br><% =txtProblemdesc %></span>
			      </td>
				  </tr>
				  <tr>
				    <td nowrap valign="top" style="padding-top:3;">Failure Reason: &nbsp;</td>
				    <td valign="top">
						  <input type="text" name="txtFailure" id="txtFailure" value="<% =txtfailure %>" <%If WO_CLOSE_SHOWFAILUREANALYSIS_FREQ = "Yes" Then%>class="required1" <%Else%>class="normal" <%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="14" onChange="top.dovalid('FA',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" fkm="FA"><img id="imgFailure" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('FA_F',txtFailure,'WO')" class="lookupicon" WIDTH="16" HEIGHT="20">
						  <span id="txtFailureDesc" class="mc_lookupdesc"><br><% =txtfailuredesc %></span>
			      </td>
				  </tr>
				  <tr>
				    <td nowrap valign="top" style="padding-top:3;">Solution: &nbsp;</td>
				    <td valign="top">
						  <input type="text" name="txtSolution" id="txtSolution" value="<% =txtSolution %>" <%If WO_CLOSE_SHOWFAILUREANALYSIS_SREQ = "Yes" Then%>class="required1" <%Else%>class="normal" <%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="14" onChange="top.dovalid('FA',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" fkm="FA"><img id="imgSolution" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('FA_S',txtSolution,'WO')" class="lookupicon" WIDTH="16" HEIGHT="20">
						  <span id="txtSolutionDesc" class="mc_lookupdesc"><br><% =txtSolutiondesc %></span>
			      </td>
				  </tr>
			  </table>
		  </td>
<!--
      <td nowrap valign="top" onclick="txtFailedWO.click();" align="right" style="cursor:hand;">
        <input <% 'If txtfailureWO Then %>checked <% 'End If %>onclick="this.blur();" type="checkbox" value="ON" name="txtFailedWO" tabindex="">Failed
        Work Order
      </td>
	  </tr>
-->
      <td align="right" valign="top">
        <table border="0" cellpadding="0" cellspacing="0" style="font-family:Arial; font-size:9pt; color:#000000">
          <tr>
            <td nowrap valign="top" onclick="txtFailedWO.click();" align="" style="cursor:hand;"><input <% If txtfailureWO Then %>checked <% End If %>onclick="this.blur();" type="checkbox" value="ON" name="txtFailedWO" tabindex="8">Failed Work Order</td>
          </tr>

          <tr>
            <td nowrap valign="top" onclick="txtFollowUpAllWO.click();" align="" style="cursor:hand;"><input <% If txtFollowUpAllWO Then %>checked <% End If %>onclick="checkfollowup(document.mcform.txtFollowupChoice);this.blur();" type="checkbox" value="ON" name="txtFollowUpAllWO" tabindex="8">Create Follow-up WO(s)</td>
          </tr>

          <tr>
            <td style="padding-left:23px;">
              <select style="width:196;" class="explorerselects" tabindex="-1" size="1" name="txtFollowupChoice" onchange="checkfollowup(this);">
                <option selected value="1">Single WO for all Failed Tasks</option>
                <option value="2">Separate WOs for each Failed Task</option>
              </select>
            </td>
          </tr>

          <tr style="display:none;">
            <td nowrap valign="top" onclick="txtFollowUpSingleWO.click();" align="" style="cursor:hand;"><input <% If txtFollowUpSingleWO Then %>checked <% End If %>onclick="this.blur();" type="checkbox" value="ON" name="txtFollowUpSingleWO" tabindex="8">Create Single Follow-up<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;WO for all Failed Tasks</td>
          </tr>
          <tr style="display:none;">
            <td nowrap valign="top" onclick="txtFollowUpMultiWO.click();" align="" style="cursor:hand;"><input <% If txtFollowUpMultiWO Then %>checked <% End If %>onclick="this.blur();" type="checkbox" value="ON" name="txtFollowUpMultiWO" tabindex="8">Create Multiple Follow-up<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;WOs for each Failed Task</td>
          </tr>
        </table>
      </td>


<!--
      <td align="right" valign="top">
        <table border="0" cellpadding="0" cellspacing="0" style="font-family:Arial; font-size:9pt; color:#000000">
          <tr>
            <td nowrap valign="top" onclick="txtFailedWO.click();" align="" style="cursor:hand;">
            <input <% 'If txtfailureWO Then %>checked <% 'End If %>onclick="this.blur();" type="checkbox" value="ON" name="txtFailedWO" tabindex="8">
            Failed Work Order</td>
          </tr>
          <tr>
            <td nowrap valign="top" onclick="txtFollowUpSingleWO.click();" align="" style="cursor:hand;">
            <input <% 'If txtFollowUpSingleWO Then %>checked <% 'End If %>onclick="this.blur();" type="checkbox" value="ON" name="txtFollowUpSingleWO" tabindex="8">
            Create Single Follow-up<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;WO for all Failed Tasks</td>
          </tr>
          <tr>
            <td nowrap valign="top" onclick="txtFollowUpMultiWO.click();" align="" style="cursor:hand;">
            <input <% 'If txtFollowUpMultiWO Then %>checked <% 'End If %>onclick="this.blur();" type="checkbox" value="ON" name="txtFollowUpMultiWO" tabindex="8">
            Create Multiple Follow-up<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;WOs for each Failed Task</td>
          </tr>
        </table>
      </td>
-->

    </tr>
  </table>
<%End Sub

'Show Labor Hours
Sub WO_CLOSE_SHOWLABORHOURS
%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-family:Arial; font-size:9pt; color:#000000">
		<tr>
			<td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
				<img SRC="../../images/icons/laborsm_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="15">Labor Hours
  		</td>
	  </tr>
	  <tr>
		  <td height="5">
		  </td>
	  </tr>
	  <tr>
		  <td onclick="txtMyLaborHrs.click();" style="cursor:hand;" width="50%">
        Set My Total Labor Hours to: <input mcType="N" class="normalright" type="text" name="txtMyLaborHrs" size="1" style="width:32" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.fnTrapAlpha(this);" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" value="" tabindex=""><img id="imgtxtMyLaborHrs" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtMyLaborHrs);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
      </td>
    </tr>
	</table>
<%End Sub

'Show Meter Readings
Sub WO_CLOSE_SHOWMETERREADINGS
%>
  <div id="option_meters" style="<% If mode = "WO" and assetexists Then %>display:;<% Else %>display:none;<% End If %>">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-family:Arial; font-size:9pt; color:#000000">
      <tr>
        <td colspan="2" class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
          &nbsp;<img SRC="../../images/icons/equip3_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="15">Meter Readings
			  </td>
		  </tr>
      <tr>
			  <td height="5"></td>
		  </tr>
      <tr>
			  <td valign="top">
				  <table border="0" cellpadding="0" cellspacing="0" style="font-family:Arial; font-size:9pt; color:#000000">
					  <tr>
					    <td nowrap valign="top" style="padding-top:3;">Meter 1 Reading: &nbsp;</td>
						  <td valign="top">
							  <input mcType="N" <%If WO_CLOSE_SHOWMETERREADINGS_REQM1 = "Yes" Then%>class="requiredright1" <%Else%>class="normalright" <%End If%> type="text" id="txtMeter1Reading" name="txtMeter1Reading" value="<% =txtmeter1reading %>" size="5" onchange="validateMeterReading(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" tabindex=""><img id="imgMeter1Reading" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtMeter1Reading);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
						  </td>
					  </tr>
				  </table>
			  </td>
			  <td valign="top" align="right">
				  <table border="0" cellpadding="0" cellspacing="0" style="font-family:Arial; font-size:9pt; color:#000000">
					  <tr>
					    <td nowrap valign="top" style="padding-top:3;">Meter 2 Reading: &nbsp;</td>
						  <td valign="top">
					 		  <input mcType="N" <%If WO_CLOSE_SHOWMETERREADINGS_REQM2 = "Yes" Then%>class="requiredright1" <%Else%>class="normalright" <%End If%> type="text" id="txtMeter2Reading" name="txtMeter2Reading" value="<% =txtmeter2reading %>" size="5" onchange="validateMeterReading(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" tabindex=""><img id="imgMeter2Reading" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtMeter2Reading);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
					    </td>
					  </tr>
				  </table>
			  </td>
		  </tr>
	  </table>
  </div>
<%End Sub

'Show Account/Category
Sub WO_CLOSE_SHOWACCOUNTCATEGORY
%>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-family:Arial; font-size:9pt; color:#000000">
	  <tr>
		  <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;" colspan="2">
	      <img SRC="../../images/icons/supportxp_g.gif" style="margin-right:5;" WIDTH="16" HEIGHT="15">Account / Category
		  </td>
	  </tr>
    <tr>
		  <td colspan="2" height="5"></td>
	  </tr>
    <tr>
		  <td colspan="2" valign="top">
			  <table width="100%" border="0" cellpadding="0" cellspacing="0" style="font-family:Arial; font-size:9pt; color:#000000">
				  <tr>
	  		    <td nowrap valign="top" style="width:60px; padding-top:3;">Account: &nbsp;</td>
				    <td valign="top">
						  <input type="text" name="txtAccount" id="txtAccount" value="<% =txtaccount %>" <%If WO_CLOSE_SHOWACCOUNTCATEGORY_AREQ = "Yes" Then%>class="required1" <%Else%>class="normal"<%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="20" onChange="top.dovalid('AC',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);top.enabledisable();" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" fkm="AC"><img id="imgAccount" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('AC',txtAccount,'WO')" class="lookupicon" WIDTH="16" HEIGHT="20">
						  <span id="txtAccountDesc" class="mc_lookupdesc"><br><% =txtaccountdesc %></span>
			      </td>
					  <td nowrap style="padding-top:0px;padding-left:16px;cursor:hand;" valign="top" onclick="txtAccountAll.click();">
						  <input onclick="this.blur();" type="checkbox" value="ON" name="txtAccountAll" tabindex="">
						  Set All
				    </td>
					  <td id="option_chargeable" style="cursor:hand;" nowrap valign="top" onclick="txtChargeable.click();" align="right">
						  <input <% If txtchargeable Then %>checked <% End If %>onclick="this.blur();" type="checkbox" value="ON" name="txtChargeable" tabindex="">
					    Chargeable
					  </td>
				  </tr>
			  </table>
		  </td>
	  </tr>
    <tr>
		  <td colspan="2" valign="top" style="padding-top:5px;">
			  <table width="75%" border="0" cellpadding="0" cellspacing="0" style="font-family:Arial; font-size:9pt; color:#000000">
				  <tr>
				    <td nowrap valign="top" style="width:60px;padding-top:3;">Category: &nbsp;</td>
					  <td valign="top">
						  <input type="text" name="txtCategory" id="txtCategory" value="<% =txtcategory %>" <%If WO_CLOSE_SHOWACCOUNTCATEGORY_CREQ = "Yes" Then%>class="required1" <%Else%>class="normal" <%End If%> mcType="C" maxlength="25" mcRequired="N" tabindex="" size="20" onChange="top.dovalid('CA',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);top.enabledisable();" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" fkm="CA_WO"><img id="imgCategory" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('CA',txtCategory,'WO')" class="lookupicon" WIDTH="16" HEIGHT="20">
						  <span id="txtCategoryDesc" class="mc_lookupdesc"><br><% =txtcategorydesc %></span>
					  </td>
					  <td nowrap style="padding-top:0px;padding-left:16px;cursor:hand;" valign="top" onclick="txtCategoryAll.click();">
			        <input onclick="this.blur();" type="checkbox" value="ON" name="txtCategoryAll" tabindex="">
					     Set All
					  </td>
				  </tr>
			  </table>
		  </td>
	  </tr>
  </table>
<%End Sub

'show Labor Actuals
Sub WO_CLOSE_SHOWLABORACTUAL
%>
<%If WO_CLOSE_SHOWLABORACTUAL_REQ = "Yes" Then%>
<div style="border: solid 1px royalblue; padding-top:5px;padding-bottom:5px;padding-left:2px;">
<%End If%>
							<!-- Labor (Actuals) Header Table -->
							<table id="ola3header" style="margin-top:0;" border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
										<img SRC="../../images/icons/laborsm_g.gif" style="margin-right:5;" width="16" height="15">Labor (Actuals)
									</td>
									<td align="right" class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
										&nbsp;<img class="mcbutton" border="0" src="../../images/button_add.gif" onclick="startPopUp('EMP'); event.cancelBubble=true; this.blur();" WIDTH="80" HEIGHT="15">
										&nbsp;<img class="mcbutton" border="0" src="../../images/button_remove.gif" onclick="mydeleterow(ola3);event.cancelBubble = true;this.blur();" WIDTH="80" HEIGHT="15">
									</td>
								</tr>
							</table>
							<%'If (Not mode = "WO") OR (Not WOGroupPK = "" and Not WOGroupPK = "-1") Then%>
							<table id="tblSplitLabor" style="margin-top:0;background-color:Yellow; <%If (Not mode = "WO") Then%>display;<% Else %>display:none;<% End If %>" border="0" cellpadding="0" cellspacing="0" width="100%">
							  <tr>
							    <td width="10" height="20" style="padding-right:8px;padding-left:6px;"><input type="checkbox" id="disperseLabor" name="disperseLabor" class="mccheckbox" <%If WO_CLOSE_SPLITLABORHOURS_CBDEFAULT = "Yes" Then Response.Write "checked" End If%> /></td>
							    <td class="fieldsheader" style="color:Black;">Split Labor Hours?</td>
							  </tr>
							</table>
							<%'End If %>
							<!-- Labor (Actuals) Header Table End -->

							<!-- Labor (Actuals) Table -->
							<table id="ola3" modid="LA" style="margin-top:0;" border="1" cellspacing="0" cellpadding="1" width="100%" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC">
								<colgroup>
									  <col span="6" height="5">
								</colgroup>
								<tbody id="ola3body" mcindex="0">
								<tr>
									<td class="tableheadercell" valign="top" width="10" height="20">
										<img alt="Click to select all" hspace="4" style="cursor:hand;" src="../../images/checkbox.gif" border="0" width="13" height="11" onclick="select_all('la3')">
									</td>
									<td mc_t="WOlabor" mc_f="WorkDate" class="tableheadercellleft" valign="top">Date</td>
									<td mc_t="WOlabor" mc_f="LaborID" class="tableheadercellleft" valign="top">Labor</td>
									<td mc_t="WOlabor" mc_f="RegularHours" class="tableheadercell" valign="top">Reg&nbsp;Hours</td>
									<td mc_t="WOlabor" mc_f="OvertimeHours" class="tableheadercell" valign="top">OT&nbsp;Hours</td>
									<td mc_t="WOlabor" mc_f="OtherHours" class="tableheadercell" valign="top">Other&nbsp;Hours</td>
								</tr>
								<tr STYLE="display: none">
									<td class="tablecboxcell" valign="top">
										<img class="lookupicon" src="../../images/undo2.gif" onclick="top.mcTRUndo(top.findRow(this));" WIDTH="14" HEIGHT="16">
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
										<input disabled class="required" mcType="D" maxlength="10" mcRequired="Y" type="text" name="txtla3WorkDateROW_ID" size="10" onChange="top.fieldvalid(this);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img id="imgla2WorkDateROW_ID" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calendar','Calendar',172,160,this,txtla2WorkDateROW_ID)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
										<span id="txtla3WorkDateROW_IDErr" class="mc_lookupdesc"></span>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
										<input disabled class="normalright" mcType="NR2" maxlength="12" mcRequired="N" type="text" name="txtla3RegularHoursROW_ID" size="1" onChange="top.fieldvalid(this);top.calclaborrow(top.findTBody(this));" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.fnTrapAlpha(this);"><img id="imgla3RegularHoursROW_ID" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calculator','Calculator',125,100,this,txtla3RegularHoursROW_ID)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
										<span id="txtla3RegularHoursROW_IDErr" class="mc_lookupdesc"></span><font style="padding-left:10;color:gray;font-weight:normal;">(Use + and - keys)</font>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
										<input disabled class="normalright" mcType="NR2" maxlength="12" mcRequired="N" type="text" name="txtla3OvertimeHoursROW_ID" size="1" onChange="top.fieldvalid(this);top.calclaborrow(top.findTBody(this));" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.fnTrapAlpha(this);"><img id="imgla3OvertimeHoursROW_ID" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calculator','Calculator',125,100,this,txtla3OvertimeHoursROW_ID)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
										<span id="txtla3OvertimeHoursROW_IDErr" class="mc_lookupdesc"></span>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
										<input disabled class="normalright" mcType="NR2" maxlength="12" mcRequired="N" type="text" name="txtla3OtherHoursROW_ID" size="1" onChange="top.fieldvalid(this);top.calclaborrow(top.findTBody(this));" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.fnTrapAlpha(this);"><img id="imgla3OtherHoursROW_ID" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calculator','Calculator',125,100,this,txtla3OtherHoursROW_ID)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
										<span id="txtla3OtherHoursROW_IDErr" class="mc_lookupdesc"></span>
									</td>
								</tr>
								<tr viewtemplate="Y" oncontextmenu="top.mcTR_OnContextMenu(this);return false;" STYLE="display: none" onmouseover="top.mcTR_OnMouseOver(this);" onmouseout="top.mcTR_OnMouseOut(this);">
									<td class="tablecboxcell" valign="top">
										<input type="checkbox" name="checkedDataList_la3ROW_ID" class="mccheckbox" onclick="top.mctr_checked(this);">
									</td>
									<td class="tabledatacellleft" valign="top">
									</td>
									<td class="tabledatacellleft" valign="top" width="140">
									</td>
									<td class="tabledatacell" nowrap valign="top">
									</td>
									<td class="tabledatacell" nowrap valign="top">
									</td>
									<td class="tabledatacell" nowrap valign="top">
									</td>
								</tr>
								<tr viewtemplate="Y" oncontextmenu="top.mcTR_OnContextMenu(this);return false;" STYLE="display: none" onclick="" onmouseover="top.mcTR_OnMouseOver(this);" onmouseout="top.mcTR_OnMouseOut(this);">
									<td class="tablecboxcell" valign="top">&nbsp;
									</td>
									<td class="tabledatacellleft" valign="top">
									</td>
									<td class="tabledatacellleft" valign="top" width="140">
									</td>
									<td class="tabledatacell" nowrap valign="top">
									</td>
									<td class="tabledatacell" nowrap valign="top">
									</td>
									<td class="tabledatacell" nowrap valign="top">
									</td>
								</tr>

								</tbody>
							</table>
							<!-- Labor (Actuals) Table End -->


<DIV  class="popup" id="pWinEmp" style="position: absolute;left:50%;top: 50%; margin-top: -200px; margin-left: -250; z-index:1000;">
<!------------------------------------------------------------------------------------------->
      <table width="100%" border="0" cellpadding="0" cellspacing="0">

        <tr>
          <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;"><img src="../../images/icons/member_attach_g.gif" />&nbsp;Add Labor (Actuals)</td>
          <td style="border-bottom:1 solid #A2A2A2;" align="right" valign="top"></td>
        </tr>

        <tr>
          <td colspan="2" style="padding-top: 4px;">
            <table width="100%" border="0" cellpadding="1" cellspacing="0">
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Labor Type</td>
                <td style="margin-top:5px;">
                  <select id="latype" tabindex="" style="font-family:arial; font-size:8pt; font-weight:bold; color:gray; width:116px;" class="explorerselects">
                    <option value="CON">Contractor</option>
                    <option value="EMP">Labor</option>
                  </select>
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Labor</td>
                <td style="margin-top:5px;">
                  <!-- This would replace both the Labot Type select and Labor lookup -->
                  <!--<input class="normal" mcType="C" maxlength="25" mcRequired="N" type="text" name="txtLabor" id="txtLabor" tabindex="19" size="10" onChange="top.dovalid('LA',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" ondblclick="top.fielddblclick(this);" oncontextmenu="top.FK_OnContextMenu(this);return false;" fkm="LA"><img id="imgLabor" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('LA',txtLabor,'WO')" class="lookupicon" WIDTH="16" HEIGHT="20">-->
                  
                  <select id="labor" tabindex="" style="font-family:arial; font-size:8pt; font-weight:bold; color:gray; width:116px;" class="explorerselects">

                  </select>
                  
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Work Date</td>
                <td style="margin-top:5px;">
                  <input ondblclick="GetCurrentDate(this);" tabindex="" mcType="D" mcName="Work Date" mcRequired="Y" type="text" name="txtWorkDate" id="txtWorkDate" value="" class="normal" size="15" maxlength="12" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"><img id="imgWorkDate" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calendar','Calendar',172,160,this,txtWorkDate)" class="lookupicon" WIDTH="16" HEIGHT="20">
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Regular Hours</td>
                <td style="margin-top:5px;">
                  <input tabindex="" mcType="N" class="normalright" type="text" id="txtHoursRegular" name="txtHoursRegular" value="" size="15" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" /><img id="imgHoursRegular" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtHoursRegular);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Overtime Hours</td>
                <td style="margin-top:5px;">
                  <input tabindex="" mcType="N" class="normalright" type="text" id="txtHoursOvertime" name="txtHoursOvertime" value="" size="15" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" /><img id="imgHoursOvertime" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtHoursOvertime);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:4px; padding-right: 10">Other Hours</td>
                <td style="padding-bottom:4px; margin-top:5px;">
                  <input tabindex="" mcType="N" class="normalright" type="text" id="txtHoursOther" name="txtHoursOther" value="" size="15" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" /><img id="imgHoursOther" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtHoursOther);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Comments</td>
                <td style="margin-top:5px;">
                  <textarea tabindex="" id="txtLaborComments" name="txtLaborComments" class="normal" cols="50" rows="2" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" ></textarea>
                </td>
              </tr>
              <tr>
                <td colspan="2" align="right" style="border-top:1 solid #A2A2A2; padding-top:4px;"><img tabindex="" class="mcbutton" border="0" src="../../images/buttonaction_apply.gif" onclick="addLaborRecord(); event.cancelBubble = true;this.blur();" WIDTH="80" HEIGHT="19">&nbsp;<img src="../../images/buttondivider.jpg" HEIGHT="19" />&nbsp;<img class="mcbutton" border="0" src="../../images/buttonaction_cancel.gif" onclick="hidePopUp('EMP');" width="80" HEIGHT="19" /></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
<!---------------------------------------------------------------------------------------------------------->
</div>

<%If WO_CLOSE_SHOWLABORACTUAL_REQ = "Yes" Then%>
</div>
<%End If %>
<%
End Sub

'Show Material Actuals
Sub WO_CLOSE_SHOWPARTACTUAL
%>
<%If WO_CLOSE_SHOWPARTACTUAL_REQ = "Yes" Then%>
<div style="border: solid 1px royalblue; padding-top:5px;padding-bottom:5px;padding-left:2px;">
<%End If%>
							<!-- Materials (Actuals) Header Table -->
							<%If Not mode = "WO" Then%>
							<table id="oma3header" style="margin-top:0; visibility:hidden;" border="0" cellpadding="0" cellspacing="0" width="100%">
							<%Else%>
							<table id="oma3header" style="margin-top:0;" border="0" cellpadding="0" cellspacing="0" width="100%">
							<%End If%>
								<tr>
									<td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
										<img src="../../images/icons/box3d_g.gif" style="margin-right:5;" width="16" height="13" alt="" />Materials
									</td>
									<td align="right" class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
										&nbsp;<img class="mcbutton" border="0" src="../../images/button_add.gif" onclick="startPopUp('INV');event.cancelBubble = true;this.blur();" width="80" height="15" alt="" />
										&nbsp;<img class="mcbutton" border="0" src="../../images/button_remove.gif" onclick="mydeleterow(oma3);event.cancelBubble = true;this.blur();" width="80" height="15" alt="" />
									</td>
								</tr>
							</table>
							<!-- Materials (Actuals) Header Table End -->

							<!-- Materials (Actuals) Table -->
							<%If Not mode = "WO" Then%>
							<table id="oma3" modid="IN" style="margin-top:0; visibility:hidden;" border="1" cellspacing="0" cellpadding="1" width="100%" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC">
							<%Else%>
							<table id="oma3" modid="IN" style="margin-top:0;" border="1" cellspacing="0" cellpadding="1" width="100%" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC">
							<%End If%>
								<colgroup>
									  <col span="4" height="5" />
								</colgroup>
								<tbody id="oma3body" mcindex="0">
								<tr>
									<td class="tableheadercell" valign="top" width="10" height="20">
										<img alt="Click to select all" hspace="4" style="cursor:hand;" src="../../images/checkbox.gif" border="0" width="13" height="11" onclick="select_all('ma3')" />
									</td>
									<!--<td mc_t="WOpart" mc_f="PartID" class="tableheadercellleft" valign="top">Item #</td>-->
									<td mc_t="WOpart" mc_f="PartName" class="tableheadercellleft" valign="top">Item Name</td>
									<td mc_t="WOpart" mc_f="LocationID" class="tableheadercellleft" valign="top">Location</td>
									<td mc_t="WOpart" mc_f="QuantityActual" class="tableheadercell" valign="top" nowrap>Est Qty</td>
									<td mc_t="WOpart" mc_f="DeliverTo" class="tableheadercellleft" valign="top" nowrap>Act Qty</td>
									<td mc_t="WOpart" mc_f="Expedite" class="tableheadercell" valign="top">Expedite</td>
									<td mc_t="WOpart" mc_f="OrderItem" class="tableheadercell" valign="top" nowrap>Order Item</td>
								</tr>
								<tr style="display: none">
									<td class="tablecboxcell" valign="top">
										<img class="lookupicon" src="../../images/undo2.gif" onclick="top.mcTRUndo(top.findRow(this));" WIDTH="14" HEIGHT="16">
									</td>
									<!--<td class="tabledatacellleft" valign="top" nowrap>
									</td>-->
									<td class="tabledatacellleft" valign="top" nowrap>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
									<td class="tabledatacellLeft" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
								</tr>
								<tr style="display: none">
									<td class="tablecboxcell" valign="top">
										<input type="checkbox" name="checkedDataList_ma3ROW_ID" class="mccheckbox" onclick="top.mctr_checked(this);">
									</td>
									<!--<td class="tabledatacellleft" valign="top" nowrap>
									</td>-->
									<td class="tabledatacellleft" valign="top" nowrap>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>									
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
								</tr>
								<tr viewtemplate="Y" oncontextmenu="top.mcTR_OnContextMenu(this);return false;" style="display: none;" onmouseover="top.mcTR_OnMouseOver(this);" onmouseout="top.mcTR_OnMouseOut(this);" >
									<td class="tablecboxcell" valign="top">&nbsp;
									</td>
									<!--<td class="tabledatacellleft" valign="top">
									</td>-->
									<td class="tabledatacellleft" valign="top">
									</td>
									<td class="tabledatacellleft" valign="top">
									</td>
									<td class="tabledatacell" nowrap valign="top">
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
									<td align="center" valign="top" nowrap>
									</td>
									<td align="center" valign="top" nowrap>
									</td>
								</tr>
								<tr style="display: none">
									<td class="tablecboxcell" valign="top">&nbsp;
									</td>
									<!--<td class="tabledatacellleft" valign="top" nowrap>
									</td>-->
									<td class="tabledatacellleft" valign="top" nowrap>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>									
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
								</tr>
								</tbody>
							</table>
							<!-- Materials (Actuals) Table End -->

<DIV class="popup" id="pWinInv" style="position: absolute;left:50%;top: 50%; margin-top: -200px; margin-left: -250; z-index:1000;">
<!------------------------------------------------------------------------------------------->
      <table width="100%" border="0" cellpadding="0" cellspacing="0">

        <tr>
          <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;"><img src="../../images/icons/member_attach_g.gif" />&nbsp;Add Material(s)</td>
          <td style="border-bottom:1 solid #A2A2A2;" align="right" valign="top"></td>
        </tr>

        <tr>
          <td colspan="2" style="padding-top: 4px;">
            <table width="100%" border="0" cellpadding="1" cellspacing="0">
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Location</td>
                <td style="margin-top:5px;">
                  <select tabindex="" id="partlocation" mcRequired="Y" style="font-family:arial; font-size:8pt; font-weight:bold; color:gray;" class="explorerselects">

                  </select>
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Part</td>
                <td style="margin-top:5px;"><input type="hidden" id="partpk" value="" size="15" />
                  <input type="text" id="partid" class="required" tabindex="" mcRequired="Y" value="" size="40" /><span id="partidDesc" class="mc_lookupdesc"></span>
                  <div id="partContainer">
                  </div>
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Quantity</td>
                <td style="margin-top:5px;">
                  <input tabindex="" mcType="N" class="requiredright" mcRequired="Y" type="text" id="txtQuantity" name="txtQuantity" value="" size="6" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" /><img id="imgQuantity" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtQuantity);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Account</td>
                <td style="margin-top:5px;"">
                  <input class="normal" mcType="C" maxlength="25" mcRequired="N" type="text" name="account" id="account" tabindex="12" size="10" onChange="top.dovalid('AC',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" ondblclick="top.fielddblclick(this);" oncontextmenu="top.FK_OnContextMenu(this);return false;" fkm="AC"><img id="imgaccount" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('AC',account,'WO')" class="lookupicon" width="16" height="20">
                  <span mchidden="y" style="display:none;" id="accountDesc" class="mc_lookupdesc"></span>
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-right: 10">Category</td>
                <td style="margin-top:5px;">
							    <input class="normal" mcType="C" maxlength="25" mcRequired="N" type="text" name="category" id="category" tabindex="12" size="10" onChange="top.dovalid('CA',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" ondblclick="top.fielddblclick(this);" oncontextmenu="top.FK_OnContextMenu(this);return false;" fkm="CA_WO"><img id="imgcategory" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('CA',category,'WO')" class="lookupicon" width="16" height="20">
							    <span mchidden="y" style="display:none;" id="categoryDesc" class="mc_lookupdesc"></span>
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-right: 10">Delivery Location</td>
                <td style="margin-top:5px;">
							    <input class="required" mcType="C" maxlength="25" mcRequired="Y" type="text" name="deliverylocation" id="deliverylocation" tabindex="13" size="40" maxlength="50" fkm="DL_WO">
							    
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right:10">Expedite</td>
                <td nowrap valign="top" onclick="txtUdfBit2.click();" style="cursor:hand; padding: 0px; margin:0px;">
                <input onclick="this.blur();" type="checkbox" value="ON" name="txtUdfBit2" tabindex="14" style="padding: 0px; margin:0px;"></td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right:10">Order Item</td>
                <td nowrap valign="top" onclick="txtUdfBit1.click();" style="cursor:hand; padding: 0px; margin:0px;">
                <input onclick="this.blur();" type="checkbox" value="ON" name="txtUdfBit1" tabindex="15" style="padding: 0px; margin:0px;"></td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Comments</td>
                <td style="margin-top:5px;">
                  <textarea tabindex="16" id="txtPartComments" name="txtPartComments" class="normal" cols="50" rows="2" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" ></textarea>
                </td>
              </tr>
              <tr style="height:5px;"><td colspan="2"></td></tr>
              <tr>
                <td colspan="2" align="right" style="border-top:1 solid #A2A2A2; padding-top:4px;"><img tabindex="" class="mcbutton" border="0" src="../../images/buttonaction_apply.gif" onclick="addMaterialRecord(); event.cancelBubble = true;this.blur();" WIDTH="80" HEIGHT="19">&nbsp;<img src="../../images/buttondivider.jpg" HEIGHT="19" />&nbsp;<img class="mcbutton" border="0" src="../../images/buttonaction_cancel.gif" onclick="hidePopUp('INV');" width="80" HEIGHT="19" /></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
<!---------------------------------------------------------------------------------------------------------->
</div>

<div class="editpopup" id="partEditWindow" style="position: absolute;left:50%;top: 50%; margin-top: -200px; margin-left: -200">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;"><img src="../../images/icons/member_attach_g.gif" />&nbsp;Update Material Qty(s)</td>
      <td style="border-bottom:1 solid #A2A2A2;" align="right" valign="top"></td>
    </tr>
    <tr>
      <td colspan="2" style="padding-top: 4px;">
        <table width="100%" border="0" cellpadding="1" cellspacing="0">
          <tr>
            <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Location</td>
            <td style="margin-top:5px;"><div id="editPartLocation" style="font-family:Arial; font-size:9pt; color:#000000"></div></td>
          </tr>
          <tr>
            <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Part</td>
            <td style="margin-top:5px;"><div id="editPartID" style="font-family:Arial; font-size:9pt; color:#000000"></div></td>
          </tr>
          <tr>
            <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Quantity</td>
            <td style="margin-top:5px;"><input type="hidden" id="partRecordID" name="partRecordID" value="" /><input type="hidden" id="partRowID" name="partRowID" value="" />
              <input tabindex="" mcType="N" class="normalright" type="text" id="editPartQuantity" name="editPartQuantity" value="" size="6" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" /><img id="imgeditPartQuantity" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,editPartQuantity);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
            </td>
          </tr>
          <tr style="height:5px;"><td colspan="2"></td></tr>
          <tr>
            <td colspan="2" align="right" style="border-top:1 solid #A2A2A2; padding-top:4px;"><img id="updbtn" mcid="" tabindex="" class="mcbutton" border="0" src="../../images/buttonaction_apply.gif" onclick="updatePartRecord(document.mcform.partRecordID.value,document.mcform.editPartQuantity.value,document.mcform.partRowID.value); event.cancelBubble = true;this.blur();" WIDTH="80" HEIGHT="19">&nbsp;<img src="../../images/buttondivider.jpg" HEIGHT="19" />&nbsp;<img class="mcbutton" border="0" src="../../images/buttonaction_cancel.gif" onclick="hideEditPopUp('INV');" width="80" HEIGHT="19" /></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</div>

<%If WO_CLOSE_SHOWPARTACTUAL_REQ = "Yes" Then%>
</div>
<%End If %>

<%
End Sub

'show Misc Cost Actuals
Sub WO_CLOSE_SHOWMISCCOSTACTUAL
%>
<%If WO_CLOSE_SHOWMISCCOSTACTUAL_REQ = "Yes" Then%>
<div style="border: solid 1px royalblue; padding-top:5px;padding-bottom:5px;padding-left:2px;">
<%End If%>

							<!-- Other Costs (Actuals) Header Table -->
							<%If Not mode = "WO" Then%>
							<table id="oot3header" style="margin-top:0;visibility:hidden;" border="0" cellpadding="0" cellspacing="0" width="100%">
							<%Else%>
							<table id="oot3header" style="margin-top:0;" border="0" cellpadding="0" cellspacing="0" width="100%">
							<%End If%>
								<tr>
									<td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
										<img SRC="../../images/icons/worldphotoxp_g.gif" style="margin-right:5;" width="16" height="13">Other Costs
									</td>
									<td align="right" class="fieldsheader" style="border-bottom:1 solid #A2A2A2;">
										&nbsp;<img class="mcbutton" border="0" src="../../images/button_add.gif" onclick="startPopUp('MISC');event.cancelBubble = true;this.blur();" width="80" height="15" alt="" />
										&nbsp;<img class="mcbutton" border="0" src="../../images/button_remove.gif" onclick="mydeleterow(oot3);event.cancelBubble = true;this.blur();" width="80" height="15" alt="" />
									</td>
								</tr>
							</table>
							<!-- Other Costs (Actuals) Header Table End -->

							<!-- Other Costs (Actuals) Table -->
							<%If Not mode = "WO" Then%>
							<table id="oot3" modid="OT" style="margin-top:0;visibility:hidden;" border="1" cellspacing="0" cellpadding="1" width="100%" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC">
							<%Else%>
							<table id="oot3" modid="OT" style="margin-top:0;" border="1" cellspacing="0" cellpadding="1" width="100%" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC">
							<%End If%>
								<colgroup>
									  <col span="3" height="5">
								</colgroup>
								<tbody id="oot3body" mcindex="0"">
								<tr>
									<td style="width:10px;heigth:20px;" class="tableheadercell" valign="top">
										<img alt="Click to select all" hspace="4" style="cursor:hand;" src="../../images/checkbox.gif" border="0" width="13" height="11" onclick="select_all('ot3')">
									</td>
									<td style="width:15%;" mc_t="WOmisccost" mc_f="MiscCostDate" class="tableheadercellleft" valign="top">Date</td>
									<td style="display:70%;" mc_t="WOmisccost" mc_f="MiscCostName" class="tableheadercellleft" valign="top">Name</td>
									<td style="width:15%;" mc_t="WOmisccost" mc_f="MiscCostPrice" class="tableheadercellright" valign="top" nowrap>Est Cost</td>
									<td style="width:15%;" mc_t="WOmisccost" mc_f="MiscCostPrice" class="tableheadercellright" valign="top" nowrap>Act Cost</td>
								</tr>
								<tr STYLE="display: none">
									<td class="tablecboxcell" valign="top">
										<img class="lookupicon" src="../../images/undo2.gif" onclick="top.mcTRUndo(top.findRow(this));" WIDTH="14" HEIGHT="16">
									</td>
									<td mcdefault="top.todaydate();" class="tabledatacellleft" valign="top" nowrap>
										<input disabled class="required" mcType="D" maxlength="10" mcRequired="Y" type="text" name="txtot3MiscCostDateROW_ID" size="15" onChange="top.fieldvalid(this);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img id="imgot3MiscCostDateROW_ID" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calendar','Calendar',172,160,this,txtot3MiscCostDateROW_ID)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
										<span id="txtot3MiscCostDateROW_IDErr" class="mc_lookupdesc"></span>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
										<input disabled class="normal" mcType="C" maxlength="1000" mcRequired="N" type="text" name="txtot3MiscCostNameROW_ID" size="36" onChange="top.fieldvalid(this);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);">
										<span id="txtot3MiscCostNameROW_IDErr" class="mc_lookupdesc"></span>
									</td>
									<td class="tabledatacellright" valign="top" nowrap>
										<input disabled class="required" mcType="D" maxlength="10" mcRequired="Y" type="text" name="txtot3MiscCostCostROW_ID" size="15" onChange="top.fieldvalid(this);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img id="imgot3MiscCostCostROW_ID" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calculator','Calculator',125,100,this,txtot3MiscCostCostROW_ID)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
										<span id="txtot3MiscCostCostROW_IDErr" class="mc_lookupdesc"></span>
									</td>
									<td class="tabledatacellright" valign="top" nowrap>
										<input disabled class="required" mcType="D" maxlength="10" mcRequired="Y" type="text" name="txtot3MiscCostCostROW_ID" size="15" onChange="top.fieldvalid(this);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img id="img1" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calculator','Calculator',125,100,this,txtot3MiscCostCostROW_ID)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
										<span id="Span1" class="mc_lookupdesc"></span>
									</td>
								</tr>
								<tr style="display: none">
									<td class="tablecboxcell" valign="top">
										<input type="checkbox" name="checkedDataList_ot3ROW_ID" class="mccheckbox" onclick="top.mctr_checked(this);">
									</td>
									<td class="tabledatacell" valign="top" nowrap>
									</td>
									<td class="tabledatacellleft" valign="top" nowrap>
									</td>
									<td class="tabledatacellright" valign="top" nowrap>
									</td>
									<td class="tabledatacellright" valign="top" nowrap>
									</td>
								</tr>
								<tr viewtemplate="Y" oncontextmenu="top.mcTR_OnContextMenu(this,false);return false;" STYLE="display: none" onclick="" onmouseover="top.mcTR_OnMouseOver(this);" onmouseout="top.mcTR_OnMouseOut(this);">
									<td class="tablecboxcell" valign="top">&nbsp;
									</td>
									<td class="tabledatacell" valign="top">
									</td>
									<td class="tabledatacellleft" valign="top">
									</td>
									<td class="tabledatacellright" valign="top">
									</td>
									<td class="tabledatacellright" valign="top">
									</td>
								</tr>


								</tbody>
							</table>
							<!-- Other Costs (Actuals) Table End -->

<DIV class="popup" id="pWinMisc" style="position: absolute;left:50%;top: 50%; margin-top: -200px; margin-left: -250; z-index:1000;">
<!------------------------------------------------------------------------------------------->
      <table width="100%" border="0" cellpadding="0" cellspacing="0">

        <tr>
          <td class="fieldsheader" style="border-bottom:1 solid #A2A2A2;"><img src="../../images/icons/worldphotoxp_g.gif" />&nbsp;Add Miscellaneous Cost(s)</td>
          <td style="border-bottom:1 solid #A2A2A2;" align="right" valign="top"></td>
        </tr>

        <tr>
          <td colspan="2" style="padding-top: 4px;">
            <table width="100%" border="0" cellpadding="1" cellspacing="0">
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Name</td>
                <td style="margin-top:5px;">
                  <input tabindex="" mcType="N" class="normal" type="text" id="txtMiscCostName" name="txtMiscCostName" value="" size="48" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" />
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Description</td>
                <td style="margin-top:5px;">
                  <textarea tabindex="" id="txtMiscCostDesc" name="txtMiscCostDesc" class="normal" cols="50" rows="2" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" ></textarea>
                </td>
              </tr>

              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Vendor</td>
                <td style="margin-top:5px;"">
						      <input class="normal" mcType="C" maxlength="25" mcRequired="N" type="text" name="MiscCostVendor" id="MiscCostVendor" tabindex="15" size="13" onChange="top.dovalid('CM',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" ondblclick="top.fielddblclick(this);" oncontextmenu="top.FK_OnContextMenu(this);return false;" fkm="CM"><img id="imgMiscCostVendor" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('CM_VENDOR',MiscCostVendor,'WO')" class="lookupicon" width="16" height="20">
						      <span id="MiscCostVendorDesc" class="mc_lookupdesc"></span>
                </td>
              </tr>

              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Date</td>
                <td style="margin-top:5px;">
                  <input tabindex="" ondblclick="GetCurrentDate(this);" mcType="D" mcName="Misc Cost Date" mcRequired="Y" type="text" name="txtMiscCostDate" id="txtMiscCostDate" value="" class="normal" size="15" maxlength="12" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onChange="top.fieldvalid(this)"><img id="imgMiscCostDate" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calendar','Calendar',172,160,this,txtMiscCostDate)" class="lookupicon" WIDTH="16" HEIGHT="20">
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Account</td>
                <td style="margin-top:5px;"">
                  <input class="normal" mcType="C" maxlength="25" mcRequired="N" type="text" name="MiscCostAccount" id="MiscCostAccount" tabindex="12" size="10" onChange="top.dovalid('AC',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" ondblclick="top.fielddblclick(this);" oncontextmenu="top.FK_OnContextMenu(this);return false;" fkm="AC"><img id="imgMiscCostAccount" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('AC',MiscCostAccount,'WO')" class="lookupicon" width="16" height="20">                  
                  <span mchidden="y" style="display:none;" id="MiscCostAccountDesc" class="mc_lookupdesc"></span>
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-right: 10">Category</td>
                <td style="margin-top:5px;">
							    <input class="normal" mcType="C" maxlength="25" mcRequired="N" type="text" name="MiscCostCategory" id="MiscCostCategory" tabindex="12" size="10" onChange="top.dovalid('CA',this,'WO');" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);" ondrop="top.dragdrop_from_list();" ondragenter="top.dragover_from_list();" ondragover="top.dragover_from_list();" ondblclick="top.fielddblclick(this);" oncontextmenu="top.FK_OnContextMenu(this);return false;" fkm="CA_WO"><img id="imgMiscCostCategory" src="../../images/lookupiconxp3fk.gif" border="0" align="absbottom" onclick="top.dolookup('CA',MiscCostCategory,'WO')" class="lookupicon" width="16" height="20">
							    <span mchidden="y" style="display:none;" id="MiscCostCategoryDesc" class="mc_lookupdesc"></span>
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Estimated Cost</td>
                <td style="margin-top:5px;">
                  <input tabindex="" mcType="N" class="normalright" type="text" id="txtPrice" name="txtPrice" value="" size="15" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" /><img id="imgPrice" src="../../images/lookupiconxp2.gif" border="0" align="absbottom" onclick="top.showpopup('calculator','Calculator',125,100,this,txtPrice);event.cancelBubble = true;" class="lookupicon" WIDTH="16" HEIGHT="20">
                </td>
              </tr>
              <tr>
                <td class="fieldlabel" style="padding-bottom:1px; padding-right: 10">Comments</td>
                <td style="margin-top:5px;">
                  <textarea tabindex="" id="txtMiscCostComments" name="txtMiscCostComments" class="normal" cols="50" rows="2" onchange="top.fieldvalid(this);" onfocus="top.fieldfocus(this)" onkeypress="return top.checkKey(this)" onblur="top.fieldblur(this)" onclick="event.cancelBubble = true;" ></textarea>
                </td>
              </tr>
              <tr style="height:5px;"><td colspan="2"></td></tr>
              <tr>
                <td colspan="2" align="right" style="border-top:1 solid #A2A2A2; padding-top:4px;"><img tabindex="" class="mcbutton" border="0" src="../../images/buttonaction_apply.gif" onclick="addMiscCostRecord(); event.cancelBubble = true;this.blur();" WIDTH="80" HEIGHT="19">&nbsp;<img src="../../images/buttondivider.jpg" HEIGHT="19" />&nbsp;<img class="mcbutton" border="0" src="../../images/buttonaction_cancel.gif" onclick="hidePopUp('MISC');" WIDTH="80" HEIGHT="19" /></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
<!---------------------------------------------------------------------------------------------------------->
</div>
<%If WO_CLOSE_SHOWMISCCOSTACTUAL_REQ = "Yes" Then%>
</div>
<%End If %>

<%
End Sub

'Sub writes existing cost actuals to their tables
Sub writerecorddata
  Set db = New ADOHelper
  sql = "SELECT PK, LaborPK, LaborID, LaborName, RegularHours, OvertimeHours, OtherHours, WorkDate, RowVersionDate FROM WOLabor WITH (NOLOCK) WHERE WOLabor.WOPK = " & WOPK & " AND WOLabor.RecordType = 2 ORDER BY WorkDate"
  Set rs = db.RunSQLReturnRS(sql,"")

  If Not db.dok Then
		errormessage = "There was a problem accessing the Work Order record. Please contact your maintenance manager for support.<br><br>" & db.derror
	Else
		If rs.eof Then
			Response.Write("<script language=""javascript"">")+nl
			Response.Write("	top.cleartable(ola3);")+nl
			Response.Write("  document.getElementById('ola3').style.display='none';")+nl
			Response.Write("</script>")+nl

		Else
			Response.Write(nl)
			Response.Write("<script language=""javascript"">")+nl
			Response.Write("  //var myframe=document.getElementById('ola3');")+nl
			Response.Write("	// Build Actual Labor Rows")+nl
			Response.Write("	// -------------------------------------------------------------------------")+nl
			Response.Write("	top.cleartable(ola3);")+nl
			Do Until rs.eof
				Response.Write("	top.builddatarow(ola3body,3,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("LaborPK") & "','WO',false,null,null,null,'" & DateNullCheck(RS("WorkDate")) & "','" & JSEncode(RS("LaborName")) & "','" & NullCheck(RS("RegularHours")) & "','" & NullCheck(RS("OvertimeHours")) & "','" & NullCheck(RS("OtherHours")) & "');")+nl
				rs.MoveNext
			Loop
			Response.Write("</script>")+nl
    End If
  End If

  sql = "SELECT WOPart.PK, WOPart.PartPK, WOPart.PartID, WOPart.PartName, WOPart.LocationPK, WOPart.LocationID, WOPart.LocationName, WOPart.QuantityEstimated, WOPart.QuantityActual," _
   & " WOPart.IssueUnitCost, WOPart.RowVersionDate, WOPart.DirectIssue, WOPart.OutOfPocket, ISNULL(est.UDFBit1,0) UDFBit1, ISNULL(est.UDFBit2, 0) UDFBit2, WOPart.RecordType" _
   & " FROM WOPart WITH (NOLOCK)" _
   & " LEFT OUTER JOIN WOPart est WITH (NOLOCK) ON WOPart.EstimatePK=est.PK" _
   & " WHERE WOPart.WOPK = " & WOPK & " AND WOPart.RecordType = 2 ORDER BY WOPart.RowVersionDate"
  Set rs = db.RunSQLReturnRS(sql,"")

  If Not db.dok Then
		errormessage = "There was a problem accessing the Work Order record. Please contact your maintenance manager for support.<br><br>" & db.derror
	Else
		If rs.eof Then
			Response.Write("<script language=""javascript"">")+nl
			Response.Write("	top.cleartable(oma3);")+nl
			Response.Write("  document.getElementById('oma3').style.display='none';")+nl
			Response.Write("</script>")+nl

		Else
			Response.Write(nl)
			Response.Write("<script language=""javascript"">")+nl
			Response.Write("	// Build Actual Part Rows")+nl
			Response.Write("	// -------------------------------------------------------------------------")+nl
			Response.Write("	top.cleartable(oma3);")+nl
         Dim locname, expd, orditem, rowtype
			Do Until rs.eof
            locname=""
            If rs("DirectIssue") then
             locname = "Direct Issue"
            ElseIf rs("OutOfPocket") then
             locname = "Out Of Pocket"
            Else
             locname = NullCheck(rs("LocationName"))
            End If
            expd="<img src=""" & Application("web_path") & Application("mapp_path") & "images/checkbox_notchecked.jpg"" border=""0"" >&nbsp;"          '"images/checkbox_notchecked.jpg"
            if BitNullCheck(rs("UDFBit2")) then 
                expd="<img src=""" & Application("web_path") & Application("mapp_path") & "images/checkbox_checked.jpg"" border=""0"">"          '"images/checkbox_checked.jpg" 
            end if
            orditem="<img src=""" & Application("web_path") & Application("mapp_path") & "images/checkbox_notchecked.jpg"" border=""0"" >&nbsp;"      '"images/checkbox_notchecked.jpg"
            if BitNullCheck(rs("UDFBit1")) then 
                orditem="<img src=""" & Application("web_path") & Application("mapp_path") & "images/checkbox_checked.jpg"" border=""0"">"   '"images/checkbox_checked.jpg" 
            end if
            rowtype = 3    
    
			Response.Write("	top.builddatarow(oma3body," &rowtype& ",null,'" & NullCheck(rs("PK")) + "$" + NullCheck(rs("RowVersionDate")) & "','" & rs("PartPK") & "','IN',false,null,null,null,'" & JSEncode(rs("PartName")) &"','" & locname & "','" & NullCheck(rs("QuantityEstimated")) & "','" & NullCheck(rs("QuantityActual")) &"','" & expd &"','" & orditem &"');")+nl
			rs.MoveNext
		Loop
		Response.Write("</script>")+nl
    End If
  End If

  'Misc Cost Items
  sql = "SELECT PK, MiscCostName, MiscCostDesc, MiscCostDate, EstimatedCost, ActualCost FROM WOMiscCost WITH (NOLOCK) WHERE WOMiscCost.WOPK = " & WOPK & " AND WOMiscCost.RecordType = 2 ORDER BY MiscCostDate"
  Set rs = db.RunSQLReturnRS(sql,"")

  If Not db.dok Then
		errormessage = "There was a problem accessing the Work Order record. Please contact your maintenance manager for support.<br><br>" & db.derror
	Else
		If rs.eof Then
			Response.Write("<script language=""javascript"">")+nl
			Response.Write("	top.cleartable(oot3);")+nl
			Response.Write("  document.getElementById('oot3').style.display='none';")+nl
			Response.Write("</script>")+nl

		Else
			Response.Write(nl)
			Response.Write("<script language=""javascript"">")+nl
			Response.Write("	// Build Actual Part Rows")+nl
			Response.Write("	// -------------------------------------------------------------------------")+nl
			Response.Write("	top.cleartable(oot3);")+nl
			Do Until rs.eof
				Response.Write("	top.builddatarow(oot3body,3,null,'" & NullCheck(rs("PK")) + "$" + NullCheck(rs("MiscCostDate")) & "','" & rs("PK") & "','',false,null,null,null,'" & JSEncode(rs("MiscCostDate")) & "','" & JSEncode(rs("MiscCostName")) &"','"&FormatCurrency(rs("EstimatedCost"),2) &"','"&FormatCurrency(rs("ActualCost"),2)&"');")+nl
				rs.MoveNext
			Loop
			Response.Write("</script>")+nl
        End If
    End If
    
    'Followup WO
    dim chkWoPk, hiWoPk
    hiWoPk = 0
    sql = "SELECT WO.WOPK, WO.WOID, WO.TargetDate, WO.RowVersionDate, lts.CodeIcon AS StatusIcon" _
            & " FROM WO WITH (NOLOCK) INNER JOIN" _
            & " LookupTableValues lts WITH (NOLOCK) ON WO.Status = lts.CodeName AND lts.LookupTable = 'WOSTATUS'" _
            & " WHERE FollowupFromWOPK = " & WOPK _
            & " ORDER BY TargetDate DESC, WOPK ASC"

    Set rs = db.RunSQLReturnRS(sql,"")

    If Not db.dok Then
		errormessage = "There was a problem accessing the Work Order record. Please contact your maintenance manager for support.<br><br>" & db.derror
    Else
		If rs.eof Then
			Response.Write("<script language=""javascript"">")+nl
			Response.Write("	top.cleartable(ofw1);")+nl
			Response.Write("  document.getElementById('ofw1').style.display='none';")+nl
			Response.Write("</script>")+nl

		Else
			Response.Write(nl)
			Response.Write("<script language=""javascript"">")+nl			
			Response.Write("	// Write Followup Work Order Rows")+nl
            Response.Write("	// -------------------------------------------------------------------------")+nl
            Response.Write("	top.cleartable(ofw1);")+nl
            Do Until rs.eof
                chkWoPk = NullCheck(RS("WOPK"))
	            Response.Write("	top.builddatarow(ofw1body,2,null,'" & chkWoPk + "$" + NullCheck(RS("RowVersionDate")) & "','" & chkWoPk & "','WO',false,'',null,null,'" & ShowImage(NullCheck(RS("StatusIcon"))) & "','" & JSEncode(RS("WOID")) & "','" & DateNullCheck(RS("TargetDate")) & "');")+nl
	            if chkWoPk > hiWoPk then hiWoPk = chkWoPk end if
	            rs.MoveNext
            Loop
            Response.Write("document.mcform.txtLastWOPK.value='" & hiWoPk & "';") 
			Response.Write("</script>")+nl
        End If
    End If  
    
    Response.Write(nl)
    
End Sub


Sub GetData
  ' Get Task Completeion
  If WO_CLOSE_ALLTASKSCOMPLETE_REQ = "Yes" Then
   
		If mode = "WOGROUP" Then
			If Not Trim(Request.Form("txtWOGroupAll")) = "" and _
			   Not WOGroupPK = "-1" and _
			   Not WOGroupPK = "" Then
				sql = "Select TaskNo from WO WITH (NOLOCK) INNER JOIN WOTask ON WOTask.WOPK = WO.WOPK WHERE (WOGroupPK = " & WOGroupPK & " OR (WOGroupPK IN (SELECT WOGroupPK FROM WO WITH (NOLOCK) WHERE WOGroupPK = " & WOGroupPK & " AND WOGroupPK > 0 AND WOGroupType = 'M'))) AND WO.IsOpen = 1 AND (WOTask.Complete = 0 AND WOTask.Fail = 0 AND WOTask.Header = 0) ORDER BY WO.WOPK"
			Else
				sql = "Select TaskNo from WO WITH (NOLOCK) INNER JOIN WOTask ON WOTask.WOPK = WO.WOPK " & Replace(actionwhere,"WHERE ","WHERE ( ") & " OR (WOGroupPK IN (SELECT WOGroupPK FROM WO WITH (NOLOCK) " & actionwhere & " AND WOGroupPK > 0 AND WOGroupType = 'M'))) AND WO.IsOpen = 1 AND (WOTask.Complete = 0 AND WOTask.Fail = 0 AND WOTask.Header = 0) ORDER BY WO.WOPK"
			End If


    Else

      sql = "SELECT TaskNo FROM WOTask WHERE WOPK = " & WOPK & " AND (Complete = 0 AND Fail = 0 AND Header = 0)"

    End If
    'Response.Write "<textarea rows=3 cols=100>" & sql & "</textarea>"
    Set rsTasks = db.RunSQLReturnRS(sql,"")

    If db.dok Then
      If Not rsTasks.EOF Then
        If mode = "WOGROUP" Then
          TaskMessage = "There are incomplete tasks on one or more of the selected work orders."
        Else
          TaskMessage = "Incomplete tasks exist on this work order."
        End If
        TaskCode = -1
      Else
        sql = "SELECT COUNT(PK) FROM WOTask WHERE WOPK = " & WOPK
        Set rsTaskCnt = db.RunSQLReturnRS(sql,"")

        If db.dok Then
          If Not rsTaskCnt.EOF Then
            If rsTaskCnt(0) > 0 Then
              TaskMessage = "All tasks are complete."
              TaskCode = 1
            Else
              If mode = "WOGROUP" Then
                TaskMessage = "There are no tasks for these work orders."
              Else
                TaskMessage = "There are no tasks for this work order."
              End If
              TaskCode = 0
            End If
          End If
        End If
      End If
	  CloseObj rsTasks
    End If

  End If
End Sub

'This is the part the shows and orders the widgets in the specified order
Sub BuildContent()
  sql = "SELECT a.PreferenceName, c.Show, Col = CASE ISNULL(a.Col,'SPANTOP') WHEN 'SPANTOP' THEN 'SPANTOP' WHEN 'SPANBOTTOM' THEN 'SPANBOTTOM' WHEN '1' THEN '1' WHEN '2' THEN '2' ELSE 'SPANTOP' END, b.Sort FROM (" & vbCRLF &_
  "	SELECT PreferenceName = REPLACE(REPLACE(P.PreferenceName,'_SORT',''),'_COL',''), Col = " & vbCRLF &_
  "		CASE " & vbCRLF &_
  "			WHEN R.preferenceValue IS NOT NULL THEN R.PreferenceValue " & vbCRLF &_
  "			ELSE P.DefaultValue	" & vbCRLF &_
  "		END " & vbCRLF &_
  "	FROM Preference P WITH (NOLOCK) LEFT OUTER JOIN RepairCenterPreference R ON P.PreferenceName = R.PreferenceName AND RepairCenterPK = " & GetSession("RCPK") & vbCRLF &_
  "	WHERE P.PreferenceCategory = 'WO_Close_EN' AND P.PreferenceName LIKE 'WO_CLOSE_SHOW%COL' " & vbCRLF &_
  ") a JOIN (	" & vbCRLF &_
  "	SELECT PreferenceName = REPLACE(REPLACE(P.PreferenceName,'_SORT',''),'_COL',''), Sort = " & vbCRLF &_
  "		CASE " & vbCRLF &_
  "			WHEN R.preferenceValue IS NOT NULL THEN R.PreferenceValue " & vbCRLF &_
  "			ELSE P.DefaultValue	" & vbCRLF &_
  "		END " & vbCRLF &_
  "	FROM Preference P WITH (NOLOCK) LEFT OUTER JOIN RepairCenterPreference R ON P.PreferenceName = R.PreferenceName AND RepairCenterPK = " & GetSession("RCPK") & vbCRLF &_
  "	WHERE P.PreferenceCategory = 'WO_Close_EN' AND P.PreferenceName LIKE 'WO_CLOSE_SHOW%SORT' " & vbCRLF &_
  ") b ON a.PreferenceName = b.PreferenceName JOIN (" & vbCRLF &_
  "	SELECT P.PreferenceName, Show = " & vbCRLF &_
  "		CASE " & vbCRLF &_
  "			WHEN R.preferenceValue IS NOT NULL THEN R.PreferenceValue " & vbCRLF &_
  "			ELSE P.DefaultValue	" & vbCRLF &_
  "		END	" & vbCRLF &_
  "	FROM Preference P WITH (NOLOCK) LEFT OUTER JOIN RepairCenterPreference R ON P.PreferenceName = R.PreferenceName AND RepairCenterPK = " & GetSession("RCPK") & vbCRLF &_
  "	WHERE P.PreferenceName NOT LIKE '%_COL' AND P.PreferenceName NOT LIKE '%_SORT' AND P.PreferenceName LIKE 'WO_CLOSE_SHOW%' " & vbCRLF &_
  ") c ON a.PreferenceName = c.PreferenceName " & vbCRLF &_
  "ORDER BY CASE a.Col WHEN 'SPANTOP' THEN 1 WHEN '1' THEN 2 WHEN '2 'THEN 3 WHEN 'SPANBOTTOM' THEN 4 END, CONVERT(INT,ISNULL(b.Sort,0))"
  Set rsa = db.RunSQLReturnRS(sql,"")
  'Response.Write "<textarea rows=6 cols=100>" & sql & "</textarea>"
  'Response.End

  If Not db.dok Then
		errormessage = "There was a problem accessing the Work Order record. Please contact your maintenance manager for support.<br><br>" & db.derror
	Else
    If Not rsa.EOF Then
      Dim subname, s, p, sc


      'Span Row - row 1
      Response.Write "<tr>"
      Response.Write "<td colspan='2'>"
      '**Start Span Content*******************************************************************************

      Do While Not rsa.EOF AND NullCheck(rsa("Col")) = "SPANTOP"
        p = NullCheck(rsa("PreferenceName"))
        sc = NullCheck(rsa("Show"))
        If sc = "No" Then
          Response.Write "<div id='"&p&"' style='display:none;'>"
        Else
          Response.Write "<div id='"&p&"' style='display:; padding-bottom:20px;'>"
        End If

        s = "Call " & p & "()"
        Execute(s)
        Response.Write "</div>"

        rsa.MoveNext
      Loop
      '**End Span Content*********************************************************************************
      Response.Write "</td>"
      Response.Write "</tr>"

      'column Row - row 2
      Response.Write "<tr>"
      'column 1
      Response.Write "<td width='50%' style='padding-right:8px;' valign='top'>"
      '**Start Column 1 Content***************************************************************************
      Do While Not rsa.EOF AND NullCheck(rsa("Col")) = "1"
        p = NullCheck(rsa("PreferenceName"))
        sc = NullCheck(rsa("Show"))
        If sc = "No" Then
          Response.Write "&nbsp;<div id='"&p&"' style='display:none; padding-bottom:20px;'>"
        Else
          Response.Write "<div id='"&p&"' style='display:; padding-bottom:20px;'>"
        End If
        s = "Call " & p & "()"
        Execute(s)
        Response.Write "</div>"

        rsa.MoveNext
      Loop
      '**End Column 1 Content*****************************************************************************
      Response.Write "</td>"
      'column 2
      Response.Write "<td valign='top' width='50%' style='padding-left:8px;'>"
      '**Start Column 2 Content***************************************************************************
      Do While Not rsa.EOF AND NullCheck(rsa("Col")) = "2"
      'Do Until rs.EOF
        p = NullCheck(rsa("PreferenceName"))
        sc = NullCheck(rsa("Show"))
        If sc = "No" Then
          Response.Write "&nbsp;<div id='"&p&"' style='display:none; padding-bottom:20px;'>"
        Else
          Response.Write "<div id='"&p&"' style='display:; padding-bottom:20px;'>"
        End If
        s = "Call " & p & "()"
        Execute(s)
        Response.Write "</div>"

        If not rsa.EOF Then
          rsa.MoveNext
        End If
      Loop
      '**End Column 2 Content*****************************************************************************
      Response.Write "</td>"
      Response.Write "</tr>"

      'Span Row - row 3
      Response.Write "<tr>"
      Response.Write "<td colspan='2'>"
      '**Start Span Content*******************************************************************************

      Do While Not rsa.EOF AND NullCheck(rsa("Col")) = "SPANBOTTOM"
        p = NullCheck(rsa("PreferenceName"))
        sc = NullCheck(rsa("Show"))
        If sc = "No" Then
          Response.Write "<div id='"&p&"' style='display:none; padding-bottom:20px;'>"
        Else
          Response.Write "<div id='"&p&"' style='display:; padding-bottom:20px;'>"
        End If

        s = "Call " & p & "()"
        Execute(s)
        Response.Write "</div>"

        rsa.MoveNext
      Loop
      '**End Span Content*********************************************************************************
      Response.Write "</td>"
      Response.Write "</tr>"
      'Close table
    End If
  End If
  CloseObj rsa
End Sub


%>
