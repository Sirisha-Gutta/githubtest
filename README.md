
## Sightly Project for Automation 


The current project is developed using Python with selenium 

You can download the latest version of Python for Windows from [Here](https://www.python.org/downloads/release/)


## Python pip modules needed

<pre><code>

python -m pip install selenium

python -m pip install webdriver-manager

python -m pip install xlrd==1.2.0

</code></pre>

## Executing tests ##


Clone Git repository https://github.com/Sirisha-Gutta/githubtest

Python Script can be run from windows command prompt  

```
example:

 c:\Users\SomeOne> python C:\MyProjectRepo\SightlyProject.py
```

or use IDE **Pycharm**

Screen "Scaling & Layout" should be set to 100%

## Scenario that is being tested ##

1. Testing was done using Chrome Browser
1. Logs into the application
1. Selects the second row to create a report as retrieving data for the first row is timing out
1. Generates report with set conditions
1. Downloads the report 
    1. reads and displays the downloaded data 
1. Compares the downloaded report with 2 hardcoded reports (correctData.xlsx & incorrectData.xlsx)
    1. The Valid data file comparision will display "Data in both sheets match"
    1. The invalid data file comparision will display the mismatched row and column number
1. Logsout of the application
