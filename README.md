# Simple Tools for OneNote

> Small toolkit for Microsoft OneNote 2013

1. Insert simple Gantt chart into Onenote page:

<img src="https://mariuskaz.github.io/images/simpleGantt.png"/>

2. Push tasks from Onenote page to [Todoist](https://todoist.com):

<img src="https://mariuskaz.github.io/images/addTasks.png"/>

References:
<ul>
	• https://support.microsoft.com/kb/2555352/en-us?wa=wsignin1.0<br/>
	• https://social.msdn.microsoft.com/Forums/office/en-US/3570a4cf-aec1-4ff7-8547-e40bf8816dd0/onenote-programming?forum=appsforoffice<br/>
	• https://code.msdn.microsoft.com/office/CSOneNoteRibbonAddIn-c3547362<br/>
	• http://msdn.microsoft.com/en-us/magazine/ff796230.aspx<br/>
	• https://github.com/Fixxzer/OneNoteRibbonAddIn<br/>
</ul>

This will demonstrate how to create a OneNote 2013 COM add-in, that implements the IRibbonExtensibility and IDTExtensibility2 interfaces, which will allow you to customize the ribbon of Microsoft OneNote 2013.

Tools required:
<ul>
	• Visual Studio 2013<br/>
	• Visual Studio Installer Projects (Available through NuGet)
	<ul>
		○ In VS: Tools -> Extensions and Updates… -> Search for "Visual Studio Installer Projects"
	</ul>
</ul>

Difficulty: This involves a good understanding of Visual Studio, OneNote, Todoist API and C#.

