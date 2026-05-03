# Bulk Assessments

Fault-tolerant C# wrappers around a script of transactions with Gemini Flash AI Models, using the Google.GenAI namespaces. Extra-slow implementation, copious use of time delays. Detailed Console logging. 

To handle quota exhaustion, it uses two round-robin queues for API Keys/Gemini Aliases and Models, respectively. Exhausts all API Keys/Alaises first, followed by Models.

## Algorithm

Upload Rubrics and Report files to the namespace's cloud URIs. 

For each Report, 
* Create a Score Workbook from an existing Template file. 
* Collect three assessments, and insert them into the Workbook. 

When complete, remove the Report and Rubrics file from the cloud.

## TODOs and Nice to Haves

* Wants to be a Windows Service
* Support for older models with higher quotas.
<br>&emsp;Getting similar results, but need to bugfix JSON issues.
* Brightspace integration
* Programmatic anonymization

## Bugs

* JSON formatting inconsistent/sketchy.
<br>&emsp;Sometimes creates schema for returned objects, especially using older models. 
* Special case exception handling 
<br>&emsp;(see code for breakpoint suggestion)