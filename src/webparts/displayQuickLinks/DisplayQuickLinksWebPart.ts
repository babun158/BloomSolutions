import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DisplayQuickLinksWebPart.module.scss';
import * as strings from 'DisplayQuickLinksWebPartStrings';
//import pnp from "sp-pnp-js";
import {
  readItems,checkUserinGroup
} from '../../commonService';
declare var $;
export interface IDisplayQuickLinksWebPartProps {
  description: string;
}

export default class DisplayQuickLinksWebPart extends BaseClientSideWebPart<IDisplayQuickLinksWebPartProps> {

  userflag: boolean = false;
  public render(): void {
    var _this = this;
    //Checking user details in group
    checkUserinGroup("Quick Links", this.context.pageContext.user.email, function (result) {
      //console.log(result);
      if (result == 1) {
        _this.userflag = true;
      }
      _this.QuickLinksDisplay();
    })
  }
  QuickLinksDisplay(){
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = 
    "<div class='head-div'>" +
    "<h3>Quick Links <a href='" + siteURL + "/Pages/ListView.aspx?CName=Quick Links' id='QuickLinkadd' style='Display:none'>More</a></h3>" +
      "</div>" +
      "<div class='emp-corner quick-link'>" +
      "<div class='carousel carousel-fade' id='carouselABC'>" +
      "<ol class='carousel-indicators' id='appendquickLinks'>" +

      "</ol>" +
      "</div>"+
      "</div>";

      this.displayQuickLinks(this.userflag);
      $('#carouselABC').carousel({ interval: 3000 });
  }
  
  public displayQuickLinks(userflag) {

    var renderhtmlcarousel = "";
    var renderhtml="";
    var count;
    let objResults;
    var activeflag = "active";    
    objResults = readItems("Quick Links", ["Title", "LinkURL", "Modified"],count, "Modified", "Display", 1);
    objResults.then((items: any[]) => {

      if (items.length > 0) {
        var carouselCount = Math.ceil(items.length / 6);
        if (carouselCount == 1) {
          //renderhtmlcarousel += "<li data-target='#carouselABC' data-slide-to='0' class='active'></li>";
        }
        else if (carouselCount == 2) {
          renderhtmlcarousel += "<li data-target='#carouselABC' data-slide-to='0' class='active'></li>";
          renderhtmlcarousel += "<li data-target='#carouselABC' data-slide-to='1' class=''></li>";
        }
        else if (carouselCount == 3) {
          renderhtmlcarousel += "<li data-target='#carouselABC' data-slide-to='0' class='active'></li>";
          renderhtmlcarousel += "<li data-target='#carouselABC' data-slide-to='1' class=''></li>";
          renderhtmlcarousel += "<li data-target='#carouselABC' data-slide-to='2' class=''></li>";
        }
        var arrItems = [];
        while (items.length) {
          arrItems.push(items.splice(0, 6));
        }
        
        renderhtml += '<div class="carousel-inner">';
        for (let index = 0; index < arrItems.length; index++) {
          if (index != 0) {
            activeflag = "";
          }
          renderhtml += '<div class="item ' + activeflag + '">';
          renderhtml += "<ul class='col-md-12'>";
          for (let innerIndex = 0; innerIndex < arrItems[index].length; innerIndex++) {
            if(arrItems[index][innerIndex].Title>18)
            {
              arrItems[index][innerIndex].Title=arrItems[index][innerIndex].Title.substring(0, 18) + "...";
            }
            if(!arrItems[index][innerIndex].LinkURL)
            {
             // arrItems[index][innerIndex]["LinkURL"]={Description: "#", Url: "#"};
            renderhtml += "<li><a href='#'>" + arrItems[index][innerIndex].Title + "</a></li>";
          }
          else
          {
            renderhtml += "<li><a href='" + arrItems[index][innerIndex].LinkURL.Url + "' target='_blank'>" + arrItems[index][innerIndex].Title + "</a></li>";
          }
          }
          renderhtml += "</ul>";
          renderhtml += "</div>";
        }
        renderhtml += "</div>";
        //renderhtml += "</div>";
        if (userflag == false) {
          $('#QuickLinkadd').hide();
        }
        else {
          $('#QuickLinkadd').show();
        }
       
      }
      else{
        if (userflag == false) {
          $('#QuickLinkadd').hide();
        }
        else {
          $('#QuickLinkadd').show();
        }
        //renderhtmlcarousel += "<li data-target='#carouselABC' data-slide-to='0' class='active'></li>";
       // renderhtml += '<div class="carousel-inner">';
       // renderhtml += '<div class="item ' + activeflag + '">';
       // renderhtml += "<ul class='col-md-12'>";
        renderhtml += "<h4 class='no-data'>No QuickLinks to display</h4>";
       // renderhtml += "</ul>";
        //  renderhtml += "</div>";
        //  renderhtml += "</div>";
         
      }
      $('#appendquickLinks').append(renderhtmlcarousel);
      $('#appendquickLinks').after(renderhtml);
    });

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
