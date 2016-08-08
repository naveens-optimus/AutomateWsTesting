# AutomateWsTesting
Tool to automate webservice testing

Sample input excel file.
For now support for excel 97-2003 (.xls).
Request tags should be in between TestCase to Run columns
Test case values should be between rows TestCase to EndTestCase
If any time, any test case should not be run, mark it "N" in the Run section
Response will be written in the Response section
Response tag for all operations should be written in the TestResult row between TestResult to EndResult
Response will printed for the result row correspondent to the Test case number.
i.e; TestCase #1, #3 and #5 has a flag “Y”, the result will be printed in the #1, #3 and #5 rows in the response section.
For now, the system executes all the Operations in the WSDL of SOAP service. Possible Request values will be taken from the Request section.

Future Enhancement planned:
1.	Control on running individual SOAP methods
2.	Support to automate REST services
3.	Support for excel xlsx.
4.	Error reporting using Email.
