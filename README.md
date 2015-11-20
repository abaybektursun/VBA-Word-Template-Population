# VBA-Word-Template-Population
Visual Basic Word macro that takes delimited data, then populates and prints a document for each record in the data file

Data File Example:                                                                                              
/////////////////////////////////////////////////                                                               
/ ID~	FName~	LName~	SuperPowers~	Skills	/                                                               
/ 12345~Mark~	Smith~	LaserBeams~		Meh     /                                                               
/////////////////////////////////////////////////                                                               

You Will Need To Change This Variables: 

...wdApp.Documents.add Template:="[YOUR TEMPLATE]"...

...With Selection.Find
            .Text = "[String You Want To Replace]"...
			
...wdApp.ActivePrinter = "[YOUR PRINTER]"...

...FileName = "[DATA FILE]"...