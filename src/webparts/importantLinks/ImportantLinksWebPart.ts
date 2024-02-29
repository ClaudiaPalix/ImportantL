import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ImportantLinksWebPart.module.scss';
import * as strings from 'ImportantLinksWebPartStrings';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IImportantLinksWebPartProps {
  description: string;
  selectedList: string; 
  Title: string;
  LinkUrl: string;
  ImageUrl: {
    Url: string;
  };
}

declare global {
  interface Window {
    chatteron: {
      open: (options: { userMessage: string }) => void;
      // Add other properties/methods if needed
    };
  }
}

export default class ImportantLinksWebPart extends BaseClientSideWebPart<IImportantLinksWebPartProps> {
  private availableLists: IPropertyPaneDropdownOption[] = [];

  protected async onInit(): Promise<void> {
    await this.userDetails();
    return super.onInit();
  }  
  
  private loadChatbot(text): void {
    try {
      if (window.chatteron && window.chatteron.open) {
        window.chatteron.open({
          userMessage: text
        });
      } else {
        console.error('Chatteron not available or open function not found.');
      }
    } catch (error) {
      console.error('Error opening chatbot:', error);
    }
  }
  
  private async loadChatterOnScript(): Promise<void> {
    const script = document.createElement('script');
    script.src = './sdk.js'; 
    script.defer = true;
    script.onload = () => {
      // Call your initialization function after the script has loaded
      this.initializeChatterOn();
    };
    document.head.appendChild(script);
  }
  
  private initializeChatterOn(): void {
    this.loadChatbot('Hello, Chatteron!');
  }

  private userEmail: string = "";

  private async userDetails(): Promise<void> {
    // Ensure that you have access to the SPHttpClient
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
  
    // Use try-catch to handle errors
    try {
      // Get the current user's information
      const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
      const userProperties: any = await response.json();
  
      console.log("User Details:", userProperties);
  
      // Access the userPrincipalName from userProperties
      const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');
  
      if (userPrincipalNameProperty) {
        this.userEmail = userPrincipalNameProperty.Value;
        console.log('User Email using User Principal Name:', this.userEmail);
      } else {
        console.error('User Principal Name not found in user properties');
      }
    } catch (error) {
      console.error('Error fetching user properties:', error);
    }
  } 


  public render(): void {
    const decodedDescription = decodeURIComponent(this.properties.description); // Decode the description (like incase there is blank space, or special characters, etc)
    console.log(decodedDescription);
    this.domElement.innerHTML = `
      <div class="${styles['important-links']}">
        <div class="${styles['parent-div']}">
        <div>
        <h2>${(decodedDescription)}</h2> <!-- Use the decoded description -->
        <div id="buttonsContainer">
        </div>
        </div>
        </div>
      </div>`;

    this._renderButtons();
  }

  private _renderButtons(): void {
    const buttonsContainer: HTMLElement | null = this.domElement.querySelector('#buttonsContainer');
    buttonsContainer?.classList.add(styles.buttonsContainer);
    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/items`;
    
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
    .then(response => response.json())
    .then(data => {
      console.log("API Response: ", data); // Log the entire response for inspection

      if (data.value && data.value.length > 0) {
        data.value.forEach((item: IImportantLinksWebPartProps) => {
          const buttonDiv: HTMLDivElement = document.createElement('div');
          buttonDiv.classList.add(styles['content-box']); // Apply the 'content-box' class styles from YourStyles.module.scss
          
          if(item.Title === "Grievances"){
            if(this.userEmail.includes("_zensar.com#EXT#")){
              buttonDiv.onclick = (event) => {
                if(item.LinkUrl.includes(siteUrl)){
                  event.preventDefault(); // Prevent the default behavior of the click event
                  const link = `${this.context.pageContext.web.absoluteUrl}/SitePages/Grievance.aspx`
                  window.location.href = link; // Navigate to the 'URL' from the API response in the same tab
                }else{
                  console.log("Grievance button clicked for external user");
                  window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/Grievance.aspx`, '_blank');
                }
                };
            }else{
            buttonDiv.onclick = async () => {
            await this.loadChatterOnScript();
              this.loadChatbot("I have a grievance to raise.");
            };
            }
          }else{
          buttonDiv.onclick = (event) => {
            if(item.LinkUrl.includes(siteUrl)){
              event.preventDefault(); // Prevent the default behavior of the click event
              window.location.href = item.LinkUrl; // Navigate to the 'URL' from the API response in the same tab
            }else{
              window.open(item.LinkUrl, '_blank'); // Open the 'Url' from the API response in a new tab
            }
          };
          }

          const imgContainer: HTMLDivElement = document.createElement('div');
          imgContainer.classList.add(styles['content-box-img-container']); // Apply the image container styles
          const img: HTMLImageElement = document.createElement('img');
          img.src = item.ImageUrl.Url; // Use the imported image URL
          imgContainer.appendChild(img); // Append the image to the container
          buttonDiv.appendChild(imgContainer); // Append the image container to the button

          const titleSpan: HTMLDivElement = document.createElement('div');
          titleSpan.classList.add(styles['content-box-text-container']);
          titleSpan.textContent = item.Title; // Use the 'Title' from the API response
          buttonDiv.appendChild(titleSpan); // Append the title to the button

          buttonsContainer!.appendChild(buttonDiv); // Append the button to the buttons container
        });
      } else {
        const noDataMessage: HTMLDivElement = document.createElement('div');
        noDataMessage.textContent = 'No applications available for the user.';
        buttonsContainer!.appendChild(noDataMessage);// Non-null assertion operator
      }
    })
    .catch(error => {
      console.error("Error fetching user data: ", error);
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadLists();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'selectedList') {
      this.setListTitle(newValue);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private _loadLists(): void {
    const listsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;
    //SPHttpClient is a class provided by Microsoft that allows developers to perform HTTP requests to SharePoint REST APIs or other endpoints within SharePoint or the host environment. It is used for making asynchronous network requests to SharePoint or other APIs in SharePoint Framework web parts, extensions, or other components.
    this.context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1)
    //SPHttpClientResponse is the response object returned after making a request using SPHttpClient. It contains information about the response, such as status code, headers, and the response body.
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: any[] }) => {
        this.availableLists = data.value.map((list) => {
          return { key: list.Title, text: list.Title };
        });
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching lists:', error);
      });
  }

  private setListTitle(selectedList: string): void {
    this.properties.selectedList = selectedList;

    this.context.propertyPane.refresh();
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
                  label: "Title For The Application"
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Select A List',
                  options: this.availableLists,
                }),
              ],
            },
          ],
        }
      ]
    };
  }
}
