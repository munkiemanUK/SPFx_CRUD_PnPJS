import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp from 'sp-pnp-js';
import styles from './CrudPnpjsWebPart.module.scss';
import * as strings from 'CrudPnpjsWebPartStrings';

export interface ICrudPnpjsWebPartProps {
  description: string;
}

export interface ISPList
{    
   ID: string;    
   Medical_Test: string;    
   Assessment_Test: string; 
}  

export default class CrudPnpjsWebPart extends BaseClientSideWebPart<ICrudPnpjsWebPartProps> {

  private AddEventListeners() : void {         
    document.getElementById('AddItemToSPList').addEventListener('click',()=>this.AddSPListItem());    
    document.getElementById('UpdateItemInSPList').addEventListener('click',()=>this.UpdateSPListItem());    
    document.getElementById('DeleteItemFromSPList').addEventListener('click',()=>this.DeleteSPListItem());    
  }    
      
  private _getSPItems(): Promise<ISPList[]> {    
    return pnp.sp.web.lists.getByTitle("Audit Tool Data").items.get().then((response) => {       
      return response;    
    });    
  }    

  private getSPItems(): void {    
    this._getSPItems()    
      .then((response) => {    
        this._renderList(response);    
      });    
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '<table class="TFtable" border=1 width=style="bordercollapse: collapse;">';
    html += `<th></th><th>ID</th><th>Assessment</th><th>Medical</th>`;
    if (items.length > 0) {
      items.forEach((item: ISPList) => {
        html += `    
          <tr>   
          <td><input type="radio" id="AuditID" name="AuditID" value="${item.ID}"><br></td>   
          <td>${item.ID}</td>    
          <td>${item.Assessment_Test}</td>    
          <td>${item.Medical_Test}</td>    
          </tr>`;
      });
    }
    else {
      html += "No records...";
    }
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#DivGetItems');
    listContainer.innerHTML = html;
  }

  public render(): void 
  {    
    this.domElement.innerHTML = `    
      <div class="parentContainer" style="background-color: white">    
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">    
         <div class="ms-Grid-col ms-u-lg ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">   
         </div>    
      </div>    
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">    
         <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:x;">CRUD Test</div>    
      </div>    
      <div style="background-color: white" >    
         <form>    
            <br>    
            <div data-role="header">    
               <h3>Add item to SharePoint List</h3>    
            </div>    
             <div data-role="main" class="ui-content">    
               <div>    
                 <input id="Assessment" placeholder="Assessment"/>    
                 <input id="Medical"  placeholder="Medical"/>    
                 <button id="AddItemToSPList"  type="submit" >Add</button>    
                 <button id="UpdateItemInSPList" type="submit" >Update</button>    
                 <button id="DeleteItemFromSPList"  type="submit" >Delete</button>  
               </div>    
             </div>    
         </form>    
      </div>    
      <br>    
      <div style="background-color: white" id="DivGetItems" />    
      </div>`;    
   this.getSPItems();    
   this.AddEventListeners();    
  }    

  AddSPListItem() {      
    pnp.sp.web.lists.getByTitle('Audit Tool Data').items.add({        
      Assessment_Test : document.getElementById('Assessment')["value"],    
      Medical_Test : document.getElementById('Medical')["value"]  
    });   
    alert("Record with Assessment type : "+ document.getElementById('Assessment')["value"] + " Added !");    
  }    

  UpdateSPListItem() {      
      var itemID =  this.domElement.querySelector('input[name = "AuditID"]:checked')["value"];  
      pnp.sp.web.lists.getByTitle("Audit Tool Data").items.getById(itemID).update({    
      Assessment_Test : document.getElementById('Assessment')["value"],    
      Medical_Test : document.getElementById('Medical')["value"]  
    });    
    alert("Record with Audit ID : "+ itemID + " Updated !");    
  }    

  DeleteSPListItem() {      
    var itemID =  this.domElement.querySelector('input[name = "AuditID"]:checked')["value"];  
    pnp.sp.web.lists.getByTitle("Audit Tool Data").items.getById(itemID).delete();    
    alert("Record with Audit ID : "+ itemID + " Deleted !");    
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
