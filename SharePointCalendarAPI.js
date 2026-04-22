/*
SharePoint Calendar API - Andrew Dagger Oct 2025
.env file not included for security reasons.
*/ 

import express, { json } from 'express';							// To create the app
import https from 'https';											// To create server
import fs from "fs";												// To import certs
import { ConfidentialClientApplication } from "@azure/msal-node";	// For authentication with microsoft graph
import asyncHandler from 'express-async-handler';					// To catch errors globally
import cors from "cors";											// Cross-Origin Resource Sharing, allows origin site to communicate with the API while keeping all other connections blocked
import dotenv from 'dotenv';										// Keeps sensitive variables from being hard-coded, stored in secrets/.env. loaded through process.env. (Not included)

// Imports sensitive variables such as credentials and urls, accessible through process.env
dotenv.config({
	path: 'secrets/.env'
});

// API credentials, used to authenticate the API with Microsoft Graph
const CLIENT_ID = process.env.CLIENT_ID;
const TENANT_ID = process.env.TENANT_ID;
const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}/`;

// Cert to be used with Microsoft Graph. This cert is installed in IIS, but the private key was exported into the API directory.
const GRAPHCERT = {
	thumbprint: process.env.THUMBPRINT,
	privateKey: fs.readFileSync("secrets/graphpk.pem")
};

// Used to create a jwt assertion token (for authentication with microsoft graph)
const cca = new ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: AUTHORITY,
    clientCertificate: GRAPHCERT
  },
});

// Allows for cross-origin resource sharing with origin site. Prevents any other site from connecting to the API.
const corsOptions = {
	origin: process.env.CORS_ORIGIN
}

// Creates the app
const app = express();
app.use(express.json(), cors(corsOptions));

// Cert used to create the https server. This cert is not in IIS. Both the cert and the private key are stored in the API directory.
// This cert and the cert used for Microsoft Graph are different.
const httpsCert = {
	key: fs.readFileSync('secrets/httpkey.pem'),
	cert: fs.readFileSync('secrets/httpcert.pem'),
};

// Starts the server.
https.createServer(httpsCert, app).listen(433, function(){
	console.log("Server running on port 443...");
});

// Listens to post requests from origin site
app.post('/', asyncHandler(async (req, res) => {

	// Log the request
	console.log(`${new Date().toISOString()}: Receieved the following request:`);
	console.log(req.body);
	console.log(`Form Sequence ID: ${req.header('Form-Sequence-ID')}.`);

	// Get the token and formDigestValue, for authentication
	const token = await getToken();
	const formDigest = await getFormDigestValue(token);

	// When a request is received, iterate over the request and post each event
	for (const e of req.body){
		await postEvent(e, token, formDigest);
	}

	// If everything succeeded, respond OK
	res.sendStatus(200);
}));

// Posts the event to the calendar in SharePoint
async function postEvent(eventData, token, formDigest) {

	console.log("Posting event...");

	// SharePoint list url
	const url = process.env.SHAREPOINT_LIST_URL;
	
	// Post the event to the sharepoint calendar
	const response = await fetch(url, {
		method: 'post',
		body: JSON.stringify(eventData),
		headers: {
			'Authorization': `Bearer ${token}`,
			'Accept': 'application/json;odata=verbose',
			'Content-Type': 'application/json',
			'Content-Length': `${JSON.stringify(eventData).length}`,
			'X-RequestDigest': `${formDigest}`
		}
	});

	// Check the response to see if the post was successful. If it wasn't, throw an exception to stop the process.
	if(!response.ok){
		throw new Error(`Error when posting event: ${JSON.stringify(eventData)}\nResponse: ${response.status} ${response.statusText}`);
	} else {
		console.log(`Success: ${JSON.stringify(eventData)}`);
	}
}

// Gets a token from Microsoft Graph
async function getToken() {
	console.log("Fetching token...");
	const result = await cca.acquireTokenByClientCredential({
		scopes: [process.env.TOKEN_SCOPE]
	});
	return result.accessToken;
}

// Gets the form digest value (a security token used to prevent cross-site request forgery)
async function getFormDigestValue(token) {
	console.log("Fetching formDigestValue...");
	// Url that contains the form digest value
	const contextInfoUrl = process.env.CONTEXT_INFO_URL;
	const response = await fetch(contextInfoUrl, {
		method: 'POST',
		headers: {
			'Authorization': `Bearer ${token}`,
			'Accept': 'application/json;odata=verbose',
			'Content-Type': 'application/json;odata=verbose'
		},
		body: '{}'
	});

	// Extract the formDigestValue from response
	const data = await response.json();
	const formDigestValue = data.d.GetContextWebInformation.FormDigestValue;
	return formDigestValue;
}

// Used to catch all exceptions globally. Responds to origin site with a 500 Internal Service Error.
// On the origin site, this prevents the form from being submitted and displays an error 
// message to the form user with the sequenceID of the form.
app.use((err, req, res, next) => {
	console.log('There was an error. Check err.txt');
	console.error(`Form Sequence ID: ${req.header('Form-Sequence-ID')}.`);
	console.error(`${err.stack}`);
	res.sendStatus(500);
});

