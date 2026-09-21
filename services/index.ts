import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export async function getDisasters(context: WebPartContext): Promise<any> {
  
  if (Math.random() < 0.5) {
    throw new Error("Simulated Api Failure");
  }

  const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists`;

  const response = await context.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
  );

  const data = await response.json();

  return data;
}
