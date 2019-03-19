import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DisplayEmployeeCornerWebPart.module.scss';
import * as strings from 'DisplayEmployeeCornerWebPartStrings';
//import pnp, { Items } from 'sp-pnp-js';
import {
  readItems,checkUserinGroup
} from '../../commonService';
declare var $;

export interface IDisplayEmployeeCornerWebPartProps {
  description: string;
}

export default class DisplayEmployeeCornerWebPart extends BaseClientSideWebPart<IDisplayEmployeeCornerWebPartProps> {

  userflag: boolean = false;
  public render(): void {
    var _this = this;
    //Checking user details in group
    checkUserinGroup("Employee Corner", this.context.pageContext.user.email, function (result) {
     // console.log(result);
      if (result == 1) {
        _this.userflag = true;
      }
      _this.QuickReadsDisplay();
    })
  }
  QuickReadsDisplay(){
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML =
      "<div class='head-div'>" +
      "<h3>Employee Corner<a href='" + siteURL + "/Pages/ListView.aspx?CName=Employee Corner' id='EmployeeAdd' style='Display:none'>More</a></h3>" +
      "</div>" +
      "<div class='emp-corner'>" +
      "<div class='carousel carousel-fade' id='carousel123'>" +
      "<ol class='carousel-indicators' id='appendEmployeeCorner'>" +

      "</ol>" +
      "</div>"+
      "</div>";
    this.displayEmployeeCorner(this.userflag);
    $('#carousel123').carousel({ interval: 5000 });
  }

  
  displayEmployeeCorner(userflag) {
    var siteURL = this.context.pageContext.web.absoluteUrl;

    var renderhtmlitems = "";
    var renderhtmlcarousel = "";
    var renderhtml = "";
    var count;
    let objResults;
    var activeflag = "active";
    
    objResults = readItems("Employee Corner", ["Title", "Modified", "Display", "DocumentFile", "FileLeafRef", "File_x0020_Type", "EncodedAbsUrl"], count, "Modified", "Display", 1);
    objResults.then((items: any[]) => {
      if (items.length > 0) {
        var carouselCount = Math.ceil(items.length / 3);
        if (carouselCount == 1) {
          //renderhtmlcarousel += "<li data-target='#carousel123' data-slide-to='0' class='active'></li>";
        }
        else if (carouselCount == 2) {
          renderhtmlcarousel += "<li data-target='#carousel123' data-slide-to='0' class='active'></li>";
          renderhtmlcarousel += "<li data-target='#carousel123' data-slide-to='1' class=''></li>";
        }
        else if (carouselCount == 3) {
          renderhtmlcarousel += "<li data-target='#carousel123' data-slide-to='0' class='active'></li>";
          renderhtmlcarousel += "<li data-target='#carousel123' data-slide-to='1' class=''></li>";
          renderhtmlcarousel += "<li data-target='#carousel123' data-slide-to='2' class=''></li>";
        }
        var arrItems = [];
        while (items.length) {
          arrItems.push(items.splice(0, 3));
        }
        renderhtml += '<div class="carousel-inner">';

        for (let index = 0; index < arrItems.length; index++) {
          if (index != 0) {
            activeflag = "";
          }
          renderhtml += '<div class="item ' + activeflag + '">';
          renderhtml += "<ul class='col-md-12'>";
          for (let innerIndex = 0; innerIndex < arrItems[index].length; innerIndex++) {
            var FileType=arrItems[index][innerIndex].DocumentFile.Url.split("/");
            FileType=FileType[FileType.length-1].split(".").pop(-1);
            if(arrItems[index][innerIndex].Title>18){
              arrItems[index][innerIndex].Title=arrItems[index][innerIndex].Title.substring(0, 18) + "...";
            }
            if (FileType == "xls" || FileType == "xlsx" ||FileType == "csv") {
              renderhtml += "<li><a href='" + arrItems[index][innerIndex].DocumentFile.Url + "'><img src='" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/xls.png'>" + arrItems[index][innerIndex].Title + "</a></li>";
            }
            else if (FileType == "pdf") {
              renderhtml += "<li><a href='" + arrItems[index][innerIndex].DocumentFile.Url + "'><img src='" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/pdf.png'>" + arrItems[index][innerIndex].Title + "</a></li>";
            } else if (FileType == "doc" || FileType == "docx") {
              renderhtml += "<li><a href='" + arrItems[index][innerIndex].DocumentFile.Url + "'><img src='" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/doc.png'>" + arrItems[index][innerIndex].Title + "</a></li>";
            } else if (FileType == "ppt") {
              renderhtml += "<li><a href='" + arrItems[index][innerIndex].DocumentFile.Url+ "'><img src='" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/ppt.png'>" + arrItems[index][innerIndex].Title + "</a></li>";
            }
          }
          renderhtml += "</ul>";
          renderhtml += "</div>";
        }
        renderhtml += "</div>";
        //renderhtml += "</div>";
        if (userflag == false) {
          $('#EmployeeAdd').hide();
        }
        else {
          $('#EmployeeAdd').show();
        }

      }
      else{
        if (userflag == false) {
          $('#EmployeeAdd').hide();
        }
        else {
          $('#EmployeeAdd').show();
        }
        //renderhtmlcarousel += "<li data-target='#carousel123' data-slide-to='0' class='active'></li>";
        // renderhtml += '<div class="carousel-inner">';
        // renderhtml += '<div class="item ' + activeflag + '">';
        // renderhtml += "<ul class='col-md-12'>";
        renderhtml += "<h4 class='no-data'>No items to display</h4>";
        // renderhtml += "</ul>";
        // renderhtml += "</div>";
        // renderhtml += "</div>";
      }
      $('#appendEmployeeCorner').append(renderhtmlcarousel);
      $('#appendEmployeeCorner').after(renderhtml);
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
