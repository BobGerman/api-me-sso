/* This code sample provides a starter kit to implement server side logic for your Teams App in TypeScript,
 * refer to https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference for complete Azure Functions
 * developer guide.
 */

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";

import repairRecords from "../repairsData.json";

import { TokenValidator, ValidateTokenOptions, getEntraJwksUri } from 'jwt-validate';

/**
 * This function handles the HTTP request and returns the repair information.
 *
 * @param {HttpRequest} req - The HTTP request.
 * @param {InvocationContext} context - The Azure Functions context object.
 * @returns {Promise<Response>} - A promise that resolves with the HTTP response containing the repair information.
 */
export async function repair(
  req: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("HTTP trigger function processed a request.");

  // Initialize response.
  const res: HttpResponseInit = {
    status: 200,
    jsonBody: {
      results: [],
    },
  };

  // Validate the access token.
  try {
    const token = req.headers.get("Authorization")?.split(" ")[1];
    if (!token) {
      throw new Error("Access token not found");
    }

    // create a new token validator for the Microsoft Entra common tenant
    const entraJwksUri = await getEntraJwksUri();
    const validator = new TokenValidator({
      jwksUri: entraJwksUri
    });

    // Use these options for single-tenant applications
    const options: ValidateTokenOptions = {
      audience: process.env["AAD_APP_CLIENT_ID"],
      issuer: `https://login.microsoftonline.com/${process.env["AAD_APP_TENANT_ID"]}/v2.0`,
      scp: ["access_as_user"]
    };

    // Use these options for multi-tenant applications
    // const options: ValidateTokenOptions = {
    //   audience: process.env["AAD_APP_CLIENT_ID"],
    //   issuer: `https://login.microsoftonline.com/${process.env["AAD_APP_TENANT_ID"]}/v2.0`,
    //   // You need to manage the list of allowed tenants on your own!
    //   // For this sample, we only allow the tenant that the app is registered in
    //   allowedTenants: [process.env["AAD_APP_TENANT_ID"]],
    //   scp: ["access_as_user"]
    // };


    // validate the token
    const validToken = await validator.validateToken(token, options);
    console.log (`Token is valid for user ${validToken.preferred_username} (${validToken.name})`);
  }
  catch (ex) {
    // Token is missing or invalid - return a 401 error
    console.error(ex);
    res.status = 401;
    res.jsonBody = {
      error: "Unauthorized",
      message: "Access token is missing or invalid"
    };
    return res;
  }

  // Get the assignedTo query parameter.
  const assignedTo = req.query.get("assignedTo");

  // If the assignedTo query parameter is not provided, return the response.
  if (!assignedTo) {
    return res;
  }

  // Filter the repair information by the assignedTo query parameter.
  const repairs = repairRecords.filter((item) => {
    const fullName = item.assignedTo.toLowerCase();
    const query = assignedTo.trim().toLowerCase();
    const [firstName, lastName] = fullName.split(" ");
    return fullName === query || firstName === query || lastName === query;
  });

  // Return filtered repair records, or an empty array if no records were found.
  res.jsonBody.results = repairs ?? [];
  return res;
}

app.http("repair", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: repair,
});
