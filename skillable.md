@lab.Title

Use this account to log into Windows:

**Username: ++@lab.VirtualMachine(Win11-Pro-Base-VM).Username++**

**Password: +++@lab.VirtualMachine(Win11-Pro-Base-VM).Password+++**

<br>

Use this account to log into Microsoft 365:

**Username: +++@lab.CloudPortalCredential(User1).Username+++**

**Password: +++@lab.CloudPortalCredential(User1).Password+++**

# Lab 445 - Build Declarative Agents for Microsoft 365 Copilot

In this lab you will build a declarative agent that assists employees of a fictitous consulting company called Trey Research. Like all declarative agents, this will use the AI model's and orchestration that's built into Microsoft 365 to provide a specialized Copilot experience that focuses on information about consultants, billing, and projects.

To make it easier, we will begin with a working declarative agent and API plugin. These are similar to what you'd get in a new project generated with Teams Toolkit, however there is a working database and sample data to work with.

The starting solution begins with access to data about consultants, but lacks general information about projects. At best it can find information about projects assigned to consultants, not projects on their own.

In the exercises that follow, you will:

 - Instruct the declarative agent on how to interact with users
 - Add a reference to a SharePoint site containing project documents
 - Add a /projects feature to the API plugin; this will show you all of the relevant packaging files needed to make the API plugin work without asking you to build the whole thing in the limited time of this lab

## Prerequisite knowledge

We assume you know the basics of creating and editing files in Visual Studio Code, and how to edit a JSON file. VS code isn't very different from other code editors, but if you have never used any code editor you might need a little help. Also JSON has a lot of squiggly braces, commas, and quotes which need to be exact in order for things to work. If you need help, please raise your hand and a proctor will explain how to compete these tasks. Thanks!

## Exercise 1: Run the starting solution

### Step 1: Open the solution in Visual Studio Code with Teams Toolkit

1. Open **Visual Studio Code**
1. Expand the **File** menu, select **Open folder...**
1. Navigate to **C:\Users\LabUser\TeamsApps**, select the folder with the name **LAB-445**, and select **Select folder**. Visual Studio Code opens the project in a new window.

Continuing in the new window:

1. On the **Activity Bar**, select the **Teams Toolkit icon** 1️⃣.
1. Under **Accounts**, select **Sign in to Microsoft 365** 2️⃣. A browser window is opened.

!IMAGE[01-04-Setup-TTK-01.png](instructions276847/01-04-Setup-TTK-01.png)

Continuing in the web browser:

- Sign in using a "Work and School" account; as a reminder here are your login credentials for the lab tenant:
    - **Username**: +++@lab.CloudPortalCredential(User1).Username+++
    - **Password**: +++@lab.CloudPortalCredential(User1).Password+++

Continuing in Visual Studio Code:

- Ensure that **Custom app upload enabled** and **Copilot access enabled** appear with a green checkbox before continuing.

!IMAGE[run-in-ttk01.png](instructions276847/run-in-ttk01.png)

### Step 2: Set up the local environment files

1. In the **Activity Bar**, select the **Explorer** (top) icon to view the files list.
1. In the env folder, rename **.env.local.user.sample** to **.env.local.user**.

### Step 3: Test the web service

- Start a new debug session, press <kbd>F5</kbd> on your keyboard.


<!-- Press F5 or hover over the "local" environment and click the debugger symbol that will be displayed 1️⃣ and then select "debug in Microsft Edge" 2️⃣.

!IMAGE[run-in-ttk02.png](instructions276847/run-in-ttk02.png) -->

It will take a while. If you get an error about not being able to run the "Ensure database" script, please try a 2nd time as this is a timing issue waiting for the Azure storage emulator to run for the first time.

The Edge browser should open to the Copilot "Bizchat" page.

If you are prompted to log in, choose "work and school" account and use these credentials:

**Username: +++@lab.CloudPortalCredential(User1).Username+++**

**Password: +++@lab.CloudPortalCredential(User1).Password+++**

Minimize the browser so you can test the API locally. (Don't close the browser or you will exit the debug session!)

With the debugger still running 1️⃣, switch to the code view in Visual Studio Code 2️⃣. Open the http folder and select the treyResearchAPI.http file 3️⃣.

Before proceeding, ensure the log file is in view by opening the "Debug console" tab 4️⃣ and ensuring that the console called "Attach to Backend" is selected 5️⃣.

Now click the "Send Request" link in treyResearchAAPI.http just above the link {{base_url}}/me 6️⃣.

!IMAGE[run-in-ttk04.png](instructions276847/run-in-ttk04.png)

You should see the response in the right panel, and a log of the request in the bottom panel. The response shows the information about the logged-in user, but since we haven't implemented authentication as yet (that's coming in Lab 6), the app will return information on the fictitious consultant "Avery Howard". Take a moment to scroll through the response to see details about Avery, including a list of project assignments.

!IMAGE[run-in-ttk05.png](instructions276847/run-in-ttk05.png)

Try some more API calls to familiarize yourself with the API and the data.

### Step 4: Run the solution in Copilot

Now restore the browser window you minimized in Step 3. You should see the Microsoft 365 Copilot window. If you need to navigate there, the URL is +++https://www.microsoft365.com/chat/?auth=2+++.

Open the right flyout 1️⃣ and, if necessary, click "Show more"2️⃣ to reveal all the choices. Then choose "Trey Genie local"3️⃣, which is the agent you just installed.

!IMAGE[run-declarative-copilot-01.png](instructions276847/run-declarative-copilot-01.png)

Try one of the prompt suggestions such as, "Find consultants with TypeScript skills." You should see two consultants, Avery Howard and Sanjay Puranik, with additional details from the database.

Your log file should reflect the request that Copilot made. You might want to try some other prompts, clicking "New Chat" in between to clear the conversation context. Here are some ideas:

 * +++Find consultants who are Azure certified and available immediately+++ (this will cause Copilot to use two query string parameters)
 * +++What projects am I assigned to?+++ (this will return information about Avery Howard who is "me" since we haven't implemented authentication)
 * +++Charge 3 hours to the Woodgrove project+++ (this will cause a POST request, and the user will need to confirm before it will udpate the data)
 * +++How many hours have I billed to Woodgrove+++ (this will demonstrate if the hours were updated in the database)

 ## Exercise 2: Add instructions and SharePoint files

 In this exercise you'll update your declarative agent with more instructions and add the capability to include knowledge from a SharePoint site.
 
### Step 1: Add instructions

Open in the **appPackage** folder open **trey-declarative-agent.json**. Add some text to the **instructions** value, staying on one line and between the quotation marks:

~~~
Be sure to remind users of the Trey motto, 'Always be Billing!'.
~~~

### Step 2: Inspect the SharePoint site

In a web browser, open the site +++https://lodsprodmca.sharepoint.com/sites/TreyLegalDocuments+++. You may need to log in again. When you see the site home page, click on "Documents" to view the Trey Research legal documents. Notice that it contains contracts for two consulting engagements, Bellows College and Woodgrove Bank.

!IMAGE[sharepoint-docs.png](instructions276847/sharepoint-docs.png)

### Step 3: Add the SharePoint capability

Now return to the **trey-declarative-agent.json** file and add these lines just above the **actions** property:

~~~
"capabilities": [
    {
        "name": "OneDriveAndSharePoint",
        "items_by_url": [
            {
                "url": "https://lodsprodmca.sharepoint.com/sites/TreyLegalDocuments"
            }
        ]
    }
],
~~~

The final **trey-declarative-agent.json** file should look like this:

~~~
{
    "$schema": "https://aka.ms/json-schemas/copilot-extensions/vNext/declarative-copilot.schema.json",
    "version": "v1.0",
    "name": "Trey Genie Local",
    "description": "You are a handy assistant for consultants at Trey Research, a boutique consultancy specializing in software development and clinical trials. ",
    "instructions": "Greet users in a professional manner, introduce yourself as the Trey Genie, and offer to help them. Your main job is to help consultants with their projects and hours. Using the TreyResearch action, you are able to find consultants based on their names, project assignments, skills, roles, and certifications. You can also find project details based on the project or client name, charge hours on a project, and add a consultant to a project. If a user asks how many hours they have billed, charged, or worked on a project, reword the request to ask how many hours they have delivered. In addition, you may offer general consulting advice. If there is any confusion, encourage users to speak with their Managing Consultant. Avoid giving legal advice. Be sure to remind users of the Trey motto, 'Always be Billing!'.",
    "conversation_starters": [
        {
            "title": "Find consultants",
            "text": "Find consultants with TypeScript skills"
        },
        {
            "title": "My Projects",
            "text": "What projects am I assigned to?"
        },
        {
            "title": "My Hours",
            "text": "How many hours have I delivered on projects this month?"
        }
    ],
    "capabilities": [
        {
            "name": "OneDriveAndSharePoint",
            "items_by_url": [
                {
                    "url": "https://lodsprodmca.sharepoint.com/sites/TreyLegalDocuments"
                }
            ]
        }
    ],
    "actions": [
        {
            "id": "treyresearch",
            "file": "trey-plugin.json"
        }
    ]
}
~~~

NOTE: The completed solution can be found in C:\Users\LabUser\TeamsApps\LAB-445-END on your workstation if you want to copy or compare with the final source code.

#### Step W: Provision a new version of the declarative agent

Let's create a new version of the declarative agent, so we can test the new capabilities.

First, in Visual Studio Code open the **env** folder and delete **.env.local** file. This will force Teams Toolkit to make a new application.

Second, in your **trey-declarative-agent.json** file, add a number to the name such as "Trey Genie 2", as you will see another copy of the agent in Copilot. Then test by clicking on the one with a new name.

### Step 4: Test in Copilot

Now press F5 or the arrow button to start the debugger again. In case BizChat doesn't open, copy the following link in the browser: +++https://www.microsoft365.com/chat/?auth=2+++.

> NOTE: If the debugger does not start after a few minutes, close Visual Studio Code and open it again. There is a race condition when starting the database a second time in the same VS Code session; it is harmless except requring restarting VS Code from time to time.

Find the "Trey Genie 2" in Copilot and test with the prompt below:


* +++"What is the status of the Woodgrove project?"+++ (The project phases should come from the statement of work in SharePoint)

In addition to including information from the statement of work, Copilot should include the Trey motto, "always be billing."

## Exercise 3: Add project API

So far, Trey Genie only knows about projects that are assigned to specific consultants. You may notice it is using the **/consultants** or **/me** paths to answer your questions. In this exercise you will add a new path to the Trey Research API, **/projects**. This will allow the declarative agent to answer more project related questions, and will give you a chance to learn about the packaging for an API plugin.

### Step 1: Add an Azure function endpoint for /projects

Create a new file, **projects.ts**, within the **/src/functions** folder, and copy this code into the new file:

~~~
/* This code sample provides a starter kit to implement server side logic for your Teams App in TypeScript,
 * refer to https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference for complete Azure Functions
 * developer guide.
 */

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import ProjectApiService from "../services/ProjectApiService";
import { ApiProject, ApiAddConsultantToProjectResponse, ErrorResult } from "../model/apiModel";
import { HttpError, cleanUpParameter } from "../services/Utilities";
import IdentityService from "../services/IdentityService";

/**
 * This function handles the HTTP request and returns the project information.
 *
 * @param {HttpRequest} req - The HTTP request.
 * @param {InvocationContext} context - The Azure Functions context object.
 * @returns {Promise<Response>} - A promise that resolves with the HTTP response containing the project information.
 */

// Define a Response interface.
interface Response extends HttpResponseInit {
    status: number;
    jsonBody: {
        results: ApiProject[] | ApiAddConsultantToProjectResponse | ErrorResult;
    };
}
export async function projects(
    req: HttpRequest,
    context: InvocationContext
): Promise<Response> {
    context.log("HTTP trigger function projects processed a request.");
    // Initialize response.
    const res: Response = {
        status: 200,
        jsonBody: {
            results: [],
        },
    };

    try {

        // Will throw an exception if the request is not valid
        const userInfo = await IdentityService.validateRequest(req);

        const id = req.params?.id?.toLowerCase();
        let body = null;
        switch (req.method) {
            case "GET": {

                let projectName = req.query.get("projectName")?.toString().toLowerCase() || "";
                let consultantName = req.query.get("consultantName")?.toString().toLowerCase() || "";

                console.log(`➡️ GET /api/projects: request for projectName=${projectName}, consultantName=${consultantName}, id=${id}`);

                projectName = cleanUpParameter("projectName", projectName);
                consultantName = cleanUpParameter("consultantName", consultantName);

                if (id) {
                    const result = await ProjectApiService.getApiProjectById(id);
                    res.jsonBody.results = [result];
                    console.log(`   ✅ GET /api/projects: response status ${res.status}; 1 projects returned`);
                    return res;
                }

                // Use current user if the project name is user_profile
                if (projectName.includes('user_profile')) {
                    const result = await ProjectApiService.getApiProjects("", userInfo.name);
                    res.jsonBody.results = result;
                    console.log(`   ✅ GET /api/projects for current user response status ${res.status}; ${result.length} projects returned`);
                    return res;
                }

                const result = await ProjectApiService.getApiProjects(projectName, consultantName);
                res.jsonBody.results = result;
                console.log(`   ✅ GET /api/projects: response status ${res.status}; ${result.length} projects returned`);
                return res;
            }
            case "POST": {
                switch (id.toLocaleLowerCase()) {
                    case "assignconsultant": {
                        try {
                            const bd = await req.text();
                            body = JSON.parse(bd);
                        } catch (error) {
                            throw new HttpError(400, `No body to process this request.`);
                        }
                        if (body) {
                            const projectName = cleanUpParameter("projectName", body["projectName"]);
                            if (!projectName) {
                                throw new HttpError(400, `Missing project name`);
                            }
                            const consultantName = cleanUpParameter("consultantName", body["consultantName"]?.toString() || "");
                            if (!consultantName) {
                                throw new HttpError(400, `Missing consultant name`);
                            }
                            const role = cleanUpParameter("Role", body["role"]);
                            if (!role) {
                                throw new HttpError(400, `Missing role`);
                            }
                            let forecast = body["forecast"];
                            if (!forecast) {
                                forecast = 0;
                                //throw new HttpError(400, `Missing forecast this month`);
                            }
                            console.log(`➡️ POST /api/projects: assignconsultant request, projectName=${projectName}, consultantName=${consultantName}, role=${role}, forecast=${forecast}`);
                            const result = await ProjectApiService.addConsultantToProject
                                (projectName, consultantName, role, forecast);

                            res.jsonBody.results = {
                                status: 200,
                                clientName: result.clientName,
                                projectName: result.projectName,
                                consultantName: result.consultantName,
                                remainingForecast: result.remainingForecast,
                                message: result.message
                            };

                            console.log(`   ✅ POST /api/projects: response status ${res.status} - ${result.message}`);
                        } else {
                            throw new HttpError(400, `Missing request body`);
                        }
                        return res;
                    }
                    default: {
                        throw new HttpError(400, `Invalid command: ${id}`);
                    }
                }

            }
            default: {
                throw new Error(`Method not allowed: ${req.method}`);
            }
        }

    } catch (error) {

        const status = <number>error.status || <number>error.response?.status || 500;
        console.log(`   ⛔ Returning error status code ${status}: ${error.message}`);

        res.status = status;
        res.jsonBody.results = {
            status: status,
            message: error.message
        };
        return res;
    }
}

app.http("projects", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    route: "projects/{*id}",
    handler: projects,
});
~~~

This will add the **/projects** requests to your Azure function using database code that was already in place.

### Step 2: Add /projects to the HTTP test file

Edit the **/http/treyResearchAPI.http** file and add these lines at the bottom of file:

~~~

########## /api/projects - working with projects ##########

### Get all projects
{{base_url}}/projects

### Get project by id
{{base_url}}/projects/1

### Get project by project or client name
{{base_url}}/projects/?projectName=supply

### Get project by consultant name
{{base_url}}/projects/?consultantName=dominique

### Add consultant to project
POST {{base_url}}/projects/assignConsultant
Content-Type: application/json

{
    "projectName": "contoso",
    "consultantName": "sanjay",
    "role": "architect",
    "forecast": 30
}
~~~

### Step 3: Add /projects to the Open API Definition

Open the file **/appPackage/trey-definition.json**. This file documents the Trey Research API using the Open API Specification (OAS) format. This is often referred to as a "Swagger" file because OAS documents used to be called Swagger files.

Let's add a new endpoint to the API specification. The code snippet below makes a **GET** request for the **/projects** path, including query string parameters for **consultantName** and **projectName**.

Find **"paths": {** array and copy these lines inside the array right after **"paths": {** line:

~~~
"/projects/": {
    "get": {
        "operationId": "getProjects",
        "summary": "Get projects matching a specified project name and/or consultant name",
        "description": "Returns detailed information about projects matching the specified project name and/or consultant name",
        "parameters": [
            {
                "name": "consultantName",
                "in": "query",
                "description": "The name of the consultant assigned to the project",
                "required": false,
                "schema": {
                    "type": "string"
                }
            },
            {
                "name": "projectName",
                "in": "query",
                "description": "The name of the project or name of the client",
                "required": false,
                "schema": {
                    "type": "string"
                }
            }
        ],
        "responses": {
            "200": {
                "description": "Successful response",
                "content": {
                    "application/json": {
                        "schema": {
                            "type": "object",
                            "properties": {
                                "results": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "name": {
                                                "type": "string"
                                            },
                                            "description": {
                                                "type": "string"
                                            },
                                            "location": {
                                                "type": "object",
                                                "properties": {
                                                    "street": {
                                                        "type": "string"
                                                    },
                                                    "city": {
                                                        "type": "string"
                                                    },
                                                    "state": {
                                                        "type": "string"
                                                    },
                                                    "country": {
                                                        "type": "string"
                                                    },
                                                    "postalCode": {
                                                        "type": "string"
                                                    },
                                                    "latitude": {
                                                        "type": "number"
                                                    },
                                                    "longitude": {
                                                        "type": "number"
                                                    }
                                                }
                                            },
                                            "mapUrl": {
                                                "type": "string",
                                                "format": "uri"
                                            },
                                            "role": {
                                                "type": "string"
                                            },
                                            "forecastThisMonth": {
                                                "type": "integer"
                                            },
                                            "forecastNextMonth": {
                                                "type": "integer"
                                            },
                                            "deliveredLastMonth": {
                                                "type": "integer"
                                            },
                                            "deliveredThisMonth": {
                                                "type": "integer"
                                            }
                                        }
                                    }
                                },
                                "status": {
                                    "type": "integer"
                                }
                            }
                        }
                    }
                }
            },
            "404": {
                "description": "Project not found"
            }
        }
    }
},
~~~

Let's add another endpoint to the API specification. The code snippet below makes a **POST** request for the **/projects/assignConsultant** path. 

Now add these lines inside the same **"paths": {** array:

~~~
"/projects/assignConsultant": {
    "post": {
        "operationId": "postAssignConsultant",
        "summary": "Assign consultant to a project when name, role and project name is specified.",
        "description": "Assign (add) consultant to a project when name, role and project name is specified.",
        "requestBody": {
            "required": true,
            "content": {
                "application/json": {
                    "schema": {
                        "type": "object",
                        "properties": {
                            "projectName": {
                                "type": "string"
                            },
                            "consultantName": {
                                "type": "string"
                            },
                            "role": {
                                "type": "string"
                            },
                            "forecast": {
                                "type": "integer"
                            }
                        },
                        "required": [
                            "projectName",
                            "consultantName",
                            "role",
                            "forecast"
                        ]
                    }
                }
            }
        },
        "responses": {
            "200": {
                "description": "Successful assignment",
                "content": {
                    "application/json": {
                        "schema": {
                            "type": "object",
                            "properties": {
                                "results": {
                                    "type": "object",
                                    "properties": {
                                        "status": {
                                            "type": "integer"
                                        },
                                        "clientName": {
                                            "type": "string"
                                        },
                                        "projectName": {
                                            "type": "string"
                                        },
                                        "consultantName": {
                                            "type": "string"
                                        },
                                        "remainingForecast": {
                                            "type": "integer"
                                        },
                                        "message": {
                                            "type": "string"
                                        }
                                    }
                                },
                                "status": {
                                    "type": "integer"
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
~~~

Be sure to check your nesting on the brackets as it gets a little tricky with large JSON files! For your reference the finished file is at **C:\Users\LabUser\TeamsApps\Lab-445-Completed\appPackage\trey-definition.json**.

### Step 4: Add the projects information to your API plugin file

The API plugin file contains additional information about your API that isn't included in the OAS (swagger) standard. Here we will add two "functions" - API functions, one for the **/projects** GET request and another for the POST.

Open your **appPackage/trey-plugin.json** file and find **"functions": [** line.

Insert the GET request function for **/projects** right after **"functions":[** line:

~~~
{
    "name": "getProjects",
    "description": "Returns detailed information about projects matching the specified project name and/or consultant name",
    "capabilities": {
    "response_semantics": {
        "data_path": "$.results",
        "properties": {
        "title": "$.name",
        "subtitle": "$.description"
        },
        "static_template": {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
            "type": "Container",
            "$data": "${$root}",
            "items": [
                {
                "speak": "${description}",
                "type": "ColumnSet",
                "columns": [
                    {
                    "type": "Column",
                    "items": [
                        {
                        "type": "TextBlock",
                        "text": "${name}",
                        "weight": "bolder",
                        "size": "extraLarge",
                        "spacing": "none",
                        "wrap": true,
                        "style": "heading"
                        },
                        {
                        "type": "TextBlock",
                        "text": "${description}",
                        "wrap": true,
                        "spacing": "none"
                        },
                        {
                        "type": "TextBlock",
                        "text": "${location.city}, ${location.country}",
                        "wrap": true
                        },
                        {
                        "type": "TextBlock",
                        "text": "${clientName}",
                        "weight": "Bolder",
                        "size": "Large",
                        "spacing": "Medium",
                        "wrap": true,
                        "maxLines": 3
                        },
                        {
                        "type": "TextBlock",
                        "text": "${clientContact}",
                        "size": "small",
                        "wrap": true
                        },
                        {
                        "type": "TextBlock",
                        "text": "${clientEmail}",
                        "size": "small",
                        "wrap": true
                        }
                    ]
                    },
                    {
                    "type": "Column",
                    "items": [
                        {
                        "type": "Image",
                        "url": "${mapUrl}",
                        "altText": "${location.street}"
                        }
                    ]
                    }
                ]
                }
            ]
            },
            {
            "type": "TextBlock",
            "text": "Project Metrics",
            "weight": "Bolder",
            "size": "Large",
            "spacing": "Medium",
            "horizontalAlignment": "Center",
            "separator": true
            },
            {
            "type": "ColumnSet",
            "columns": [
                {
                "type": "Column",
                "width": "stretch",
                "items": [
                    {
                    "type": "TextBlock",
                    "text": "Forecast This Month",
                    "weight": "Bolder",
                    "spacing": "Small",
                    "horizontalAlignment": "Center"
                    },
                    {
                    "type": "TextBlock",
                    "text": "${forecastThisMonth} ",
                    "size": "ExtraLarge",
                    "weight": "Bolder",
                    "horizontalAlignment": "Center"
                    }
                ]
                },
                {
                "type": "Column",
                "width": "stretch",
                "items": [
                    {
                    "type": "TextBlock",
                    "text": "Forecast Next Month",
                    "weight": "Bolder",
                    "spacing": "Small",
                    "horizontalAlignment": "Center"
                    },
                    {
                    "type": "TextBlock",
                    "text": "${forecastNextMonth} ",
                    "size": "ExtraLarge",
                    "weight": "Bolder",
                    "horizontalAlignment": "Center"
                    }
                ]
                }
            ]
            },
            {
            "type": "ColumnSet",
            "columns": [
                {
                "type": "Column",
                "width": "stretch",
                "items": [
                    {
                    "type": "TextBlock",
                    "text": "Delivered Last Month",
                    "weight": "Bolder",
                    "spacing": "Small",
                    "horizontalAlignment": "Center"
                    },
                    {
                    "type": "TextBlock",
                    "text": "${deliveredLastMonth} ",
                    "size": "ExtraLarge",
                    "weight": "Bolder",
                    "horizontalAlignment": "Center"
                    }
                ]
                },
                {
                "type": "Column",
                "width": "stretch",
                "items": [
                    {
                    "type": "TextBlock",
                    "text": "Delivered This Month",
                    "weight": "Bolder",
                    "spacing": "Small",
                    "horizontalAlignment": "Center"
                    },
                    {
                    "type": "TextBlock",
                    "text": "${deliveredThisMonth} ",
                    "size": "ExtraLarge",
                    "weight": "Bolder",
                    "horizontalAlignment": "Center"
                    }
                ]
                }
            ]
            }
        ],
        "actions": [
            {
            "type": "Action.OpenUrl",
            "title": "View map",
            "url": "${mapUrl}"
            }
        ]
        }
    }
    }
},
~~~

Notice that in addition to the name and description, this includes **"response_semantics"** which tell Copilot the most important parts of your API response. It also includes a **"static_template"** which is an adaptive card which data binds to the HTTP response body to display project details.

Now add another function after **"functions":[** line for the Post request function for **projects/assignConsultant**:

~~~
,
{
    "name": "postAssignConsultant",
    "description": "Assign (add) consultant to a project when name, role and project name is specified.",
    "capabilities": {
    "response_semantics": {
        "data_path": "$",
        "properties": {
        "title": "$.results.clientName",
        "subtitle": "$.results.status"
        },
        "static_template": {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {
            "type": "TextBlock",
            "text": "Project Overview",
            "weight": "Bolder",
            "size": "Large",
            "separator": true,
            "spacing": "Medium"
            },              
            {
            "type": "Container",
            "items": [
                {
                "type": "TextBlock",
                "text": "Client Name",
                "weight": "Bolder",
                "spacing": "Small"
                },
                {
                "type": "TextBlock",
                "text": "${if(results.clientName, results.clientName, 'N/A')}",
                "wrap": true
                }
            ]
            },
            {
            "type": "Container",
            "items": [
                {
                "type": "TextBlock",
                "text": "Project Name",
                "weight": "Bolder",
                "spacing": "Small"
                },
                {
                "type": "TextBlock",
                "text": "${if(results.projectName, results.projectName, 'N/A')}",
                "wrap": true
                }
            ]
            },
            {
            "type": "Container",
            "items": [
                {
                "type": "TextBlock",
                "text": "Consultant Name",
                "weight": "Bolder",
                "spacing": "Small"
                },
                {
                "type": "TextBlock",
                "text": "${if(results.consultantName, results.consultantName, 'N/A')}",
                "wrap": true
                }
            ]
            },
            {
            "type": "Container",
            "items": [
                {
                "type": "TextBlock",
                "text": "Remaining Forecast",
                "weight": "Bolder",
                "spacing": "Small"
                },
                {
                "type": "TextBlock",
                "text": "${if(results.remainingForecast, results.remainingForecast, 'N/A')}",
                "wrap": true
                }
            ]
            },
            {
            "type": "Container",
            "items": [
                {
                "type": "TextBlock",
                "text": "Message",
                "weight": "Bolder",
                "spacing": "Small"
                },
                {
                "type": "TextBlock",
                "text": "${if(results.message, results.message, 'N/A')}",
                "wrap": true
                }
            ]
            }            
        ]
        
        }
        
    },
    "confirmation": {
        "type": "AdaptiveCard",
        "title": "Assign consultant to a project when name, role and project name is specified.",
        "body": "* **ProjectName**: {{function.parameters.projectName}}\n* **ConsultantName**: {{function.parameters.consultantName}}\n* **Role**: {{function.parameters.role}}\n* **Forecast**: {{function.parameters.forecast}}"
    }
    }
}
~~~

Find **"run_for_functions": [** and update it by adding the new functions **postAssignConsultant** and **getProjects**. The final version of "run_for_functions" should look like below:

~~~
"run_for_functions": [       
     "getConsultants",        
     "getUserInformation",        
     "postBillhours",
     "postAssignConsultant",
     "getProjects"   
]
~~~


Again, please double check your nesting and commas as editing large JSON files can be tricky! The correctly modified file is on your lab workstation in **C:\Users\LabUser\TeamsApps\Lab-445-Completed\appPackage/trey-plugin.json**.

#### Step W: Provision a new version of the declarative agent

Let's create a new version of the declarative agent, so we can test the new capabilities.

First, in Visual Studio Code open the **env** folder and delete **.env.local** file. This will force Teams Toolkit to make a new application.

Second, in your **trey-declarative-agent.json** file, add a number to the name such as "Trey Genie 3", as you will see another copy of the agent in Copilot. Then test by clicking on the one with a new name.

### Step 5: Test the API

Now restart the debugger. Although the code is updated automatically, you need to completely restart it to force it to redeploy the app package, which now contains more details.

Once it has started, verify that the new API paths are working by minimizing (not closing) the browser and opening the **http/treyResearchAPI.http** file. 
This time try sending the GET request for all projects.

~~~
### Get all projects
{{base_url}}/projects
~~~

You should get back ten projects.

### Step 6: Test the updated declarative agent in Copilot

With the debugger still running, restore your debug browser session. Open Copilot and the "Trey Genie 3" declarative agent.
Here are a couple of prompts to try:

* +++What projects is Trey Resarch working on now?+++ (should return all the projects)
* +++Please add Domi as a designer on the Contoso project. Forecast 30 hours for her work.+++ (should show a confirmation card, then add Domi to the project)
* +++What projects is Domi working on?+++ (should now include the Contoso project).

# Congratulations!

---
You have completed Lab 445 and built a Declarative agent with an API plugin.
If you want to learn more, including how to add API authentication to your project, you can find a deeper dive into this and other examples at [https://aka.ms/copilotdevcamp](https://aka.ms/copilotdevcamp).

What cool prompts can you think of that weren't mentioned in the lab instructions?

