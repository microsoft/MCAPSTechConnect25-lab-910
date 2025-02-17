@lab.Title

If you're attending the lab at MCAPS Tech Connect 2025 in Seattle, you can use the following credentials.

Use this account to log into Windows:

**Username: ++@lab.VirtualMachine(Win11-Pro-Base-VM).Username++**

**Password: +++@lab.VirtualMachine(Win11-Pro-Base-VM).Password+++**

<br>

Use this account to log into Microsoft 365:

**Username: +++@lab.CloudPortalCredential(User1).Username+++**

**Password: +++@lab.CloudPortalCredential(User1).Password+++**

# Lab 910 - Build Declarative Agents for Microsoft 365 Copilot

In this lab you will build a declarative agent that assists employees of a fictitous consulting company called Trey Research. Like all declarative agents, this will use the AI model's and orchestration that's built into Microsoft 365 to provide a specialized Copilot experience that focuses on information about consultants, billing, and projects.

To make it easier, we will begin with a working declarative agent and API plugin. These are similar to what you'd get in a new project generated with Teams Toolkit, however there is a working database and sample data to work with.

The starting solution begins with access to data about consultants, but lacks general information about projects. At best it can find information about projects assigned to consultants, not projects on their own.

In the exercises that follow, you will:

 - Instruct the declarative agent on how to interact with users
 - Add a reference to a SharePoint site containing project documents
 - Add a /projects feature to the API plugin; this will show you all of the relevant packaging files needed to make the API plugin work without asking you to build the whole thing in the limited time of this lab

## Prerequisite knowledge

We assume you know the basics of creating and editing files in Visual Studio Code, and how to edit a JSON file. VS code isn't very different from other code editors, but if you have never used any code editor you might need a little help. Also JSON has a lot of squiggly braces, commas, and quotes which need to be exact in order for things to work. If you need help, please raise your hand and a proctor will explain how to compete these tasks. Thanks!