# TuntsChallenge

How to run the application

To run this application usign node.js, run <npm install googleapis@39 --save>To create a project and enable an API, refer to Create a project and enable the API
After that, create a project and enable an API on Google Cloud Platform, follow this link: https://developers.google.com/workspace/guides/create-project
To get your credentials, go to https://developers.google.com/workspace/guides/create-credentials
You need to past your credentials in "client_email" inside your credentials.json file.
Now you need to learn how to identify your sheet ID: i.e. https://docs.google.com/spreadsheets/d/HereIsYourSpreadsheetID/edit#gid=0 
Then, run this command <npm install googleapis@39 --save>
Then, npm install -g typescript to install npm typescript package
If you are using Windows, you might give permission on Powershell Execution Policies. To do that, run the command on Powershell as an Admin <Set-ExecutionPolicy -ExecutionPolicy AllSigned -Scope LocalMachine>

To compile your TS code into a JS Code, run <tsc --build tsconfig.json>
To run your JS application, run <node ./dist/alunos.js>
  
Your application should be up and running right now!


