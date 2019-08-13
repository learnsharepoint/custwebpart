import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CustWebPartWebPart.module.scss';
import * as strings from 'CustWebPartWebPartStrings';

export interface ICustWebPartWebPartProps {
  listname: string;
}

export default class CustWebPartWebPart extends BaseClientSideWebPart<ICustWebPartWebPartProps> {
  private data : any;

  private updateCustomers(custs : any[]) {
    let html = "";

    console.log("updatecustomers()-> Bulding html... ");
    custs.forEach(c => {
      html += `<div>ID: ${ c.CustomerID } <br/>
                Name : ${ c.Title } <br/>
                Web : ${ c.Email.Url } </br/>
              </div>
      `;
    });
    console.log("updateCustomers() done!");
    console.log("Built html = " + html);

    
    // Inject the HTML with div id = customerinfo
    console.log("Updating the UI with html..");
    this.domElement.querySelector('#customerinfo').innerHTML = html;
  }

  private getCustomers() : any {
    let listName : string = this.properties.listname;
    let data : any = {};

    let url = this.context.pageContext.web.absoluteUrl 
                  + "/_api/Lists/GetByTitle('" + listName + "')/Items?$orderby=CustomerID&$select=CustomerID,Title,Email";

    this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => {
        res.json().then((d) => {
          console.log("Data returned by WS Call : " + JSON.stringify(d));

          data = d;
          this.data = data;

          // Call a function to update the UI
          this.updateCustomers(data.value);
        });
      })
      .catch((err) => { 
        console.log("Error in WS Call to " + url);
        return null;
      });

    return data;
  }


  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.custWebPart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div id="customerinfo">Loading...</div>
          </div>
        </div>
      </div>`;

      // Make the Call
      console.log("Calling _api to get Customer Info...");
      this.data = this.getCustomers();
      console.log("Data returned from WS Call : " + this.data);

      // TBD - Build the HTML manually and inject into the div on line 50
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Customer WebPart Info"
          },
          groups: [
            {
              groupName: "List Info",
              groupFields: [
                PropertyPaneTextField('listname', {
                  label: "List Name"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
