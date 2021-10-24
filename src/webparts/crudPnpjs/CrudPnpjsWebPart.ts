import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './CrudPnpjsWebPart.module.scss';
import * as strings from 'CrudPnpjsWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import * as React from 'react';
import pnp from 'sp-pnp-js';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

// import node module external libraries
import * as $ from 'jquery';
require('bootstrap');

export interface ICrudPnpjsWebPartProps {
  description: string;
}

export interface AuditDataList
{    
   ID: string;    
   Medicals: string;    
   Assessment: string;
   Audit_Year: Date;
   ohaName: string;
   ctlName: string;
   Month: string;
   Case_Number: string;
   Audit_Date: Date;
}  

export interface AuditQuestionList
{    
   ID: string;    
   Section: string;    
   Assessment: string;
   Question_Number: Number;
   Question_Text: string;
   Min_Outcome: Number;
} 

export default class CrudPnpjsWebPart extends BaseClientSideWebPart<ICrudPnpjsWebPartProps> {

  private AddEventListeners() : void {         
    document.getElementById('AddItemToSPList').addEventListener('click',()=>this.AddSPListItem());    
    document.getElementById('UpdateItemInSPList').addEventListener('click',()=>this.UpdateSPListItem());    
    document.getElementById('DeleteItemFromSPList').addEventListener('click',()=>this.DeleteSPListItem());    
  }    
      
  private _getAuditData(): Promise<AuditDataList[]> {    
    return pnp.sp.web.lists.getByTitle("Audit Tool Data").items.get().then((response) => {       
      return response;    
    });    
  }    

  private getAuditData(): void {    
    this._getAuditData()    
      .then((response) => {    
        this._renderAuditData(response);    
      });    
  }

  private _renderAuditData(items: AuditDataList[]): void {
    let html: string = '<table class="TFtable" border=1 width=style="bordercollapse: collapse;">';
    html += `<th></th><th>ID</th><th>Assessment</th><th>Medical</th>`;
    if (items.length > 0) {
      items.forEach((item: AuditDataList) => {
        html += `    
          <tr>   
          <td><input type="radio" id="AuditID" name="AuditID" value="${item.ID}"><br></td>   
          <td>${item.ID}</td>    
          <td>${item.Assessment}</td>    
          <td>${item.Medicals}</td>    
          </tr>`;
      });
    }
    else {
      html += "No records...";
    }
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#AuditDataItems');
    listContainer.innerHTML = html;
  }

  private _getAuditQuestions(): Promise<AuditQuestionList[]> {    
    return pnp.sp.web.lists.getByTitle("Audit Tool Questions").items.get().then((response) => {       
      return response;    
    });    
  }    

  private getAuditQuestions(): void {    
    this._getAuditQuestions()    
      .then((response) => {    
        this._renderQuestions(response);    
      });    
  }  

  private _renderQuestions(items: AuditQuestionList[]): void {
    let html="";
    if (items.length > 0) {
      items.forEach((item: AuditQuestionList) => {
        html += `    
        <div class="row">
            <div class="col-1">${item.Question_Number}</div>
            <div class="col-6">${item.Question_Text}</div>
            <div class="col-3 text-center">
                <div class="form-check-inline">
                    <label class="form-check-label">
                    Yes <input type="radio" class="form-check-input" name="CCRQ1yes">
                    </label>
                </div>
                <div class="form-check-inline">
                    <label class="form-check-label">
                    No <input type="radio" class="form-check-input" name="CCRQ1no">
                    </label>
                </div>
                <div class="form-check-inline">
                    <label class="form-check-label">
                    N/A <input type="radio" class="form-check-input" name="CCRQ1na">
                    </label>
                </div> 
            </div>
            <div class="col bg-success text-white border border-success text-center">100%</div>
            <div class="col text-right">${item.Min_Outcome}</div>
        </div>
        <hr>
        `;
      });
    }
    else {
      html += "No records...";
    }
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#AuditQuestionItems');
    listContainer.innerHTML = html;
  }

  public render(): void {
    let bootstrapCssURL = "https://stackpath.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    let fontawesomeCssURL = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/regular.min.css";
    SPComponentLoader.loadCss(bootstrapCssURL);
    SPComponentLoader.loadCss(fontawesomeCssURL);
    
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
      <br/>    
      <div style="background-color: white" id="AuditDataItems" />    
      </div>
      <div class="row text-white" style="background-color: #545487;">
          <h3 class="ml-2">Clinical Consultation Records</h3>
      </div>
      <div class="row">
          <div class="col-10"></div>
          <div class="col">Average<br/>Yearly<br/>Score</div>
          <div class="col">Minimum<br/>Outcome</div>
      </div>
      <hr>
      <div class="container" style="overflow-y:scroll; overflow-x:hidden; height: 25vh !important;" id="AuditQuestionItems"></div>            
      `;    
   this.getAuditData(); 
   this.getAuditQuestions();   
   this.AddEventListeners();    
  }    

  public AddSPListItem() {      
    pnp.sp.web.lists.getByTitle('Audit Tool Data').items.add({        
      Assessment : document.getElementById('Assessment')["value"],    
      Medicals : document.getElementById('Medical')["value"]  
    });   
    alert("Record with Assessment type : "+ document.getElementById('Assessment')["value"] + " Added !");    
  }    

  public UpdateSPListItem() {      
      var itemID =  this.domElement.querySelector('input[name = "AuditID"]:checked')["value"];  
      pnp.sp.web.lists.getByTitle("Audit Tool Data").items.getById(itemID).update({    
      Assessment : document.getElementById('Assessment')["value"],    
      Medicals : document.getElementById('Medical')["value"]  
    });    
    alert("Record with Audit ID : "+ itemID + " Updated !");    
  }    

  public DeleteSPListItem() {      
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
