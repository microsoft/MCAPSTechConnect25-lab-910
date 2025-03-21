# Lab 910 - Build Declarative Agents for Microsoft 365 Copilot

In this lab you will build a declarative agent that assists employees of a fictitious consulting company called Trey Research. Like all declarative agents, this will use the AI model's and orchestration that's built into Microsoft 365 to provide a specialized Copilot experience that focuses on information about consultants, billing, and projects.

To make it easier, we will begin with a working declarative agent and API plugin. These are similar to what you'd get in a new project generated with Teams Toolkit, however there is a working database and sample data to work with.

The starting solution begins with access to data about consultants, but lacks general information about projects. At best it can find information about projects assigned to consultants, not projects on their own.

In the exercises that follow, you will:

 - Instruct the declarative agent on how to interact with users
 - Add a reference to a SharePoint site containing project documents
 - Add a /projects feature to the API plugin; this will show you all of the relevant packaging files needed to make the API plugin work without asking you to build the whole thing in the limited time of this lab

## Prerequisite knowledge

We assume you know the basics of creating and editing files in Visual Studio Code, and how to edit a JSON file. VS code isn't very different from other code editors, but if you have never used any code editor you might need a little help. Also JSON has a lot of squiggly braces, commas, and quotes which need to be exact in order for things to work. If you need help, please raise your hand and a proctor will explain how to compete these tasks. Thanks!

## Prerequisites tools and files

To perform this lab, you will need the following requirements:

- [Visual Studio Code](https://code.visualstudio.com/)
- [The Teams Toolkit extension for Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)
- [The REST Client extension for Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=humao.rest-client)

You will need to download the lab files from the GitHub repository first:

1. Open File Explorer
2. Pick **This PC** from the left pane and double click on the C:\ drive.
3. Right click on an empty space in the File Explorer window and choose **New** > **Folder**.
4. Name it *src*.
5. Open the browser and type in the address bar the the following URL: [https://github.com/microsoft/MCAPSTechConnect25-lab-910/archive/refs/heads/agents-lab.zip](https://github.com/microsoft/MCAPSTechConnect25-lab-910/archive/refs/heads/agents-lab.zip).
6. Download the ZIP file to your computer and extract it in the *C:\src* folder you have just created.
