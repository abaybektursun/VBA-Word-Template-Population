# VBA-Word-Template-Population
Visual Basic Word macro that takes delimited data, then populates and prints a template document for each record in the delimited data file.
You Can Locate Exapmle Template Document in The Repo

Data File Example:                                                                                              
///////////////////////////////////////////////////////////                                                               
/ ID~FName~LName~SuperPowers~SkillLevel	/                                                               
/ 12345~Mark~Smith~CuteLaserBeams~Meh /                                                               
/ 12346~John~Cake~FusRoDah~Good /                                                               
///////////////////////////////////////////////////////////                                                               

You Will Need To Change The Variables : 

' ! Specify path to the template document
docPath       = ""
' ! Specify path to the data
dataPath      = ""
' ! Specify data file name
fileName      = ""
' ! Specify template document Name
tempDocName   = ""
' ! Specify printer
printer_name  = ""
' ! These are the Strings that will be replaced in the document
fieldsArray = Split("", ",")