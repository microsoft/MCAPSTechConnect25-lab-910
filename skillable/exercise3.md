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
},
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
},
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

#### Step 4: Provision a new version of the declarative agent

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
Here are a few prompts to try:

* +++What projects is Trey Resarch working on now?+++ (should return all the projects)
* +++Please add Domi as a designer on the Contoso project. Forecast 30 hours for her work.+++ (should show a confirmation card, then add Domi to the project)
* +++What projects is Domi working on?+++ (should now include the Contoso project).

> [!Note] Like in Excercise 2, since we have deployed a new declarative agent, you will need to give consent to use the API plugin even if you have previously given consent to the original agent.

# Congratulations!

---
You have completed Lab 445 and built a Declarative agent with an API plugin.
If you want to learn more, including how to add API authentication to your project, you can find a deeper dive into this and other examples at [https://aka.ms/copilotdevcamp](https://aka.ms/copilotdevcamp).

What cool prompts can you think of that weren't mentioned in the lab instructions?