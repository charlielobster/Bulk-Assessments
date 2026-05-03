# Bulk Assessments

Fault-tolerant C# wrappers around a script of transactions with Gemini Flash AI Models, using the Google.GenAI namespaces. Extra-slow implementation, copious use of time delays. Detailed Console logging. 

To handle quota exhaustion, use two round-robin queues for API Keys/Gemini Aliases and Models, respectively. Exhaust all API Keys/Alaises first, followed by Models.

## Algorithm

Upload Rubrics and Report files to the namespace's cloud URIs. 

For each Report, 
* Create a Score Workbook from an existing Template file. 
* Collect three assessments and insert them into the Workbook. 

When complete, remove the Report or Rubrics file from the cloud.

## TODOs and Nice to Haves

* Wants to be a Windows Service
* Support for slightly older AI Model with much higher quotas.
<br>&emsp;Need to bugfix JSON issues.
* Brightspace Integration
* Programmatic report anonymization implementation

## Bugs

* JSON formatting inconsistent/sketchy.
<br>&emsp;Sometimes creates its own schema for returned objects, especially using Flash-Lite models. 
* Special case exception handling 
<br>&emsp;(see code for breakpoint suggestion)