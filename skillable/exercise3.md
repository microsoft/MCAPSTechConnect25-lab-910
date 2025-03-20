## Exercise 3: Add project API

So far, Trey Genie only knows about projects that are assigned to specific consultants. You may notice it is using the **/consultants** or **/me** paths to answer your questions. The Trey Research API, however, has an additional endpoint, called **/projects**. This will allow the declarative agent to answer more project related questions, and will give you a chance to learn about the packaging for an API plugin. Let's make the required changes to enable our declarative agent to use this endpoint.

### Step 1: Add /projects to the Open API Definition

Open the file **/appPackage/trey-definition.json**. This file documents the Trey Research API using the Open API Specification (OAS) format. This is often referred to as a "Swagger" file because OAS documents used to be called Swagger files.

Let's add a new endpoint to the API specification. The code snippet below makes a **GET** request for the **/projects** path, including query string parameters for **consultantName** and **projectName**.

```json
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
```

Now let's take another look to the API specification. The code snippet below makes a **POST** request for the **/projects/assignConsultant** path. 

```json
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
```

To incorporate these changes, open the [final version of the file](../Lab-910-END/appPackage/trey-definition.json) and copy the entire content in your local **trey-definition.json** file, replacing the entire existing content.

### Step 2: Add the projects information to your API plugin file

The API plugin file contains additional information about your API that isn't included in the OAS (swagger) standard. Here we will add two "functions" - API functions, one for the **/projects** GET request and another for the POST.

Let's take a look at the snippet:

```json
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
```

Notice that in addition to the name and description, this includes **"response_semantics"** which tell Copilot the most important parts of your API response. It also includes a **"static_template"** which is an adaptive card which data binds to the HTTP response body to display project details.

Now let's take a look at the configuration for the POST request function for **projects/assignConsultant**:

```json
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
```

Finally, the last change to apply is adding inside the `run_for_functions` collection the two new functions `postAssignConsultant` and `getProjects`.

```json
"run_for_functions": [       
     "getConsultants",        
     "getUserInformation",        
     "postBillhours",
     "postAssignConsultant",
     "getProjects"   
]
```

To incorporate these changes, open the [final version of the file](../Lab-910-END/appPackage/trey-plugin.json) and copy the entire content in your local **trey-pluing.json** file, replacing the entire existing content.

#### Step 3: Provision a new version of the declarative agent

Let's create a new version of the declarative agent, so we can test the new capabilities.

First, in Visual Studio Code open the **env** folder and delete **.env.local** file. This will force Teams Toolkit to make a new application.

Second, in your **trey-declarative-agent.json** file, add a number to the name such as "Trey Genie 3", as you will see another copy of the agent in Copilot. Then test by clicking on the one with a new name.

### Step 4: Test the API

Now restart the debugger. Although the code is updated automatically, you need to completely restart it to force it to redeploy the app package, which now contains more details.

Once it has started, verify that the new API paths are working by minimizing (not closing) the browser and opening the **http/treyResearchAPI.http** file. 
This time try sending the GET request for all projects.

```text
### Get all projects
{{base_url}}/projects
```

You should get back ten projects.

### Step 5: Test the updated declarative agent in Copilot

With the debugger still running, restore your debug browser session. Open Copilot and the "Trey Genie 3" declarative agent.
Here are a few prompts to try:

* *What projects is Trey Resarch working on now?+++ (should return all the projects)*
* *Please add Domi as a designer on the Contoso project. Forecast 30 hours for her work.* (should show a confirmation card, then add Domi to the project)
* *What projects is Domi working on? (should now include the Contoso project).*

> NOTE: Like in Excercise 2, since we have deployed a new declarative agent, you will need to give consent to use the API plugin even if you have previously given consent to the original agent.

# Congratulations!

---
You have completed the lab and built a Declarative agent with an API plugin.
If you want to learn more, including how to add API authentication to your project, you can find a deeper dive into this and other examples at [https://aka.ms/copilotdevcamp](https://aka.ms/copilotdevcamp).

What cool prompts can you think of that weren't mentioned in the lab instructions?
