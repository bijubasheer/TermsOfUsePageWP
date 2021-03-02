import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { Log } from '@microsoft/sp-core-library';
import styles from './TermsOfUsePageWpWebPart.module.scss';
import * as strings from 'TermsOfUsePageWpWebPartStrings';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/regional-settings/web";
import "@pnp/sp/site-users";

const termsListName = "Terms of Use List";
const acceptanceListName = "Terms of Use Acceptance List";
const LOG_SOURCE: string = 'TermsOfUseWebPart';

let content = "";
let title = "";
let version = "";


export interface ITermsOfUsePageWpWebPartProps {
  description: string;
}

export default class TermsOfUsePageWpWebPart extends BaseClientSideWebPart<ITermsOfUsePageWpWebPartProps> {
  
  protected onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "onInit()", this.context.serviceScope);
        
    return super.onInit().then(_ => {
        sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    Log.info(LOG_SOURCE, "render()", this.context.serviceScope);
    console.log(LOG_SOURCE + " : render()" );
    this.loadData();
  }
  private async loadData()
  {
    console.log(LOG_SOURCE + " : loadData()" );
    let userId:string = await this.GetMyUserId();
    console.log(LOG_SOURCE + ": My user id = " + userId);

    await this.GetLatestVersion();
    console.log(LOG_SOURCE + ": Latest version of terms = " + version);

    let acceptedDate = await this.GetAcceptedDate(version, userId);
    console.log(LOG_SOURCE + ": Accepted on = " + acceptedDate);

    this.domElement.innerHTML = `

      <div>
        
    <p class="${ styles.description }">${content}</p>
        <p class="${ styles.description }">
          <span><b>Terms of Use  were accepted on </b>: ${acceptedDate}</span>
        </p>
        <p class="${ styles.description }">
          <span><b>Terms of Use Version</b> : ${version}</span>
        </p>
        
      </div>`;    
  }

  private async GetLatestVersion()
  {
    await sp.web.lists.getByTitle(termsListName).items.select("Title", "TermsofUseContent", "TermsVersion")
    .top(1)
    .orderBy("TermsVersion", false)
    .get()
    .then(items =>
      {
        Log.info(LOG_SOURCE, items[0]);
        content = items[0]["TermsofUseContent"];
        title = items[0]["Title"];
        version = items[0]["TermsVersion"];
      })
      .catch((err) =>
      {
        console.error(LOG_SOURCE + " : Error in GetLatestVersion()" + err);
      });
  }

  private async GetAcceptedDate(ver: string, userId: string)
  {
    let acceptedDate:string = "";
    await sp.web.lists
    .getByTitle(acceptanceListName).items
    .select("Title", "Modified", "FieldValuesAsText/Modified")
    .top(1)
    .expand("FieldValuesAsText/Modified")
    .filter("TermsVersion eq " + ver + " and " + "AcceptedBy eq " + userId)
    .get()
    .then(items =>
      {
        acceptedDate = items[0].FieldValuesAsText["Modified"];
      }
    )
    .catch((err) =>
    {
      console.error(LOG_SOURCE + ": Error in GetAcceptedDate()" + err);
    });
    return acceptedDate;
  }

  private async GetMyUserId():Promise<string>
  {
    let userId:number = -1;
    let user = await sp.web.currentUser.get().then((u: any) => { 
      userId = u.Id;
      console.log("UserId = " + userId);
    }).catch((err) =>
    {
      console.error(LOG_SOURCE + " : Error in GetAcceptedDate()" + err);
    });
    
    return Promise.resolve(userId.toString());
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
