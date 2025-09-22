import { WebPartContext } from "@microsoft/sp-webpart-base";
import { LogHelper } from "../helpers/LogHelper";
import { MSGraphClientV3 } from "@microsoft/sp-http";

class GraphService {
  private static _context: WebPartContext;

  public static async Init(context: WebPartContext): Promise<void> {
    this._context = context;
    LogHelper.info("GraphService", "Init", "Context initialized");
  }

  public static async GetExtension(extensionName: string): Promise<any> {
    // Låt fel bubbla upp
    return await this.GET(`/me/extensions/${extensionName}`);
  }

  public static async GetPreferences(extensionName: string): Promise<any> {
    // Låt fel bubbla upp
    return await this.GET(`/me/extensions/${extensionName}`);
  }

  /** Create (POST) */
  public static async SavePreferences(userSettings: any): Promise<any> {
    // Låt fel bubbla upp
    return await this.POST(`/me/extensions`, JSON.stringify(userSettings));
  }

  /** Update (PATCH) */
  public static async UpdatePreferences(userSettings: any, extensionName: string): Promise<any> {
    // Låt fel bubbla upp (kan svara 204 No Content vid OK)
    return await this.PATCH(`/me/extensions/${extensionName}`, JSON.stringify(userSettings));
  }

  private static GET(query: string): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      this._context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): void => {
        client.api(query).get((error, response) => {
          if (error) {
            // bubbla upp status
            reject(error);
            return;
          }
          resolve(response);
        });
      });
    });
  }

  private static POST(query: string, content: string) {
    return new Promise<any>((resolve, reject) => {
      this._context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): void => {
        client.api(query).post(content, (error, response, rawResponse) => {
          if (error) {
            reject(error); // error.statusCode finns ofta här
            return;
          }
          // POST brukar ge body i response – returnera den om den finns, annars råresponsen
          resolve(response ?? rawResponse);
        });
      });
    });
  }

  private static PATCH(query: string, content: string) {
    return new Promise<any>((resolve, reject) => {
      this._context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): void => {
        client.api(query).patch(content, (error, response, rawResponse) => {
          if (error) {
            reject(error); // låt status bubbla upp (t.ex. 413)
            return;
          }
          // Vid 204 No Content har vi oftast bara rawResponse
          resolve(rawResponse ?? response);
        });
      });
    });
  }
}
export default GraphService;
