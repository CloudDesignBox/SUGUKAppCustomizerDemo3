import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
//import placeholder methods
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
//import for rest calls
import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';

import * as strings from 'Sugukdemo3ApplicationCustomizerStrings';

//load jquery with jquery cycle as dependancy
import * as jQuery from 'jquery';
require('jquery-cycle');  

//import our styles
import styles from './styles.module.scss';

const LOG_SOURCE: string = 'Sugukdemo3ApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISugukdemo3ApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/* SharePoint list interface*/
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
}


/** A Custom Action which can be run during execution of a Client Side Application */
export default class Sugukdemo3ApplicationCustomizer
  extends BaseApplicationCustomizer<ISugukdemo3ApplicationCustomizerProperties> {

    //Get Announcements list data from SharePoint REST API
    private _getListData(): Promise<ISPLists> {
      //Announcements list must exist on your site for this to work
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('Announcements')/Items?$select=Title`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    }
    // ----------------------------------------------
    // ----------------------------------------------
    // render method for announcement list
    // ----------------------------------------------
    // ----------------------------------------------
    private _renderList(items: ISPList[]): void {
      //html string
      let tempHTML: string = `<div class="${styles.cycle_slideshow}" 
      data-cycle-fx="scrollHorz" 
      data-cycle-timeout="2000"
      data-cycle-slides="> div"
      >`;
      //loop through list items
      items.forEach((item: ISPList) => {
        tempHTML += `<div class="sugukannouincementitem">` + item.Title + `</div>`;
      });
      tempHTML+=`</div>`;
      //output html
      document.getElementById(styles.sugukscrolling).innerHTML = tempHTML;
      //load slideshow
      jQuery( document ).ready(function() {
        jQuery( "." + styles.cycle_slideshow).cycle();
      });
    }
    // ----------------------------------------------
    // ----------------------------------------------
    // end of render method for announcement list
    // ----------------------------------------------
    // ----------------------------------------------

    //header placeholder
    private _headerPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // ----------------------------------------------
    // ----------------------------------------------
    // Google Analytics
    // ----------------------------------------------
    // ----------------------------------------------
    let analyticsjavascript: string = '';
    analyticsjavascript= `
      (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
      (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
      m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
      })(window,document,'script','https://www.google-analytics.com/analytics.js','ga');
      
      ga('create', 'UA-XXXXX-Y', 'auto');
      ga('send', 'pageview');
      alert("Google Analytics loaded")`;
    let head: any = document.getElementsByTagName("head")[0] || document.documentElement,
    script = document.createElement("script");
    script.type = "text/javascript";

    try {
        console.log('Append child');
        script.appendChild(document.createTextNode(analyticsjavascript));
    } 
    catch (e) {
        console.log('Append child catch');
        script.text = analyticsjavascript;
    }
    head.insertBefore(script, head.firstChild);
    head.removeChild(script);
    // ----------------------------------------------
    // ----------------------------------------------
    // End of Google Analytics code
    // ----------------------------------------------
    // ----------------------------------------------

    // Header
    // Call render method for generating the needed html elements
    this._renderPlaceHolders();

    return Promise.resolve<void>();
  }
 
  private _renderPlaceHolders(): void {
    console.log('Available placeholders: ',
    this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    // Handling the top placeholder
    if (!this._headerPlaceholder) {
      this._headerPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { //on dispose method
          });
          this._headerPlaceholder.domElement.innerHTML = `
          <div class="${styles.myheader}" id="${styles.myheader}">
            <div  class="${styles.suguknewsleftpanel}">
              <div class="${styles.date}" id="${styles.date}">
              </div>
              <div class="${styles.newstitle}"">
              SUGUK Latest News:
              </div>
            </div>
            <div  class="${styles.sugukscrolling}" id="${styles.sugukscrolling}">
            </div>
          </div>
          `;
    }

    //render date and time
    this.startTime();

    //get list data for header and then render
    this._getListData()
    .then((response) => {
      this._renderList(response.value);
    });
  }

  //function to set time and date in header
  private async startTime() {
      var monthNames = ["January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
      ];
      var today = new Date();
      var day = today.getDate();
      var month = monthNames[today.getMonth()];
      var year = today.getFullYear();
      var h = today.getHours();
      var m = today.getMinutes();
      let mins :string = m.toString();
      if (m < 10) {mins = "0" + m.toString();}
      document.getElementById(styles.date).innerHTML = day.toString() + " " + month + " " + year + " " + h + ":" + mins;
}

}
//example test url
// https://clouddesignboxlimited.sharepoint.com/sites/Communication/SitePages/Home.aspx?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"b4143eed-c20e-4868-80ce-db73dcba6722":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"testMessage":"Hello as property!"}}}
//
// There must be an announcement list on the site for this to work unless you hardcode URL in REST call
