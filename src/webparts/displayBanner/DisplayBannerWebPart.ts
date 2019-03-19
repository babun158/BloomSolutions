import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
//import pnp from 'sp-pnp-js'
import styles from './DisplayBannerWebPart.module.scss';
import * as strings from 'DisplayBannerWebPartStrings';
import {
  readItems, checkUserinGroup
} from '../../commonService';
export interface IDisplayBannerWebPartProps {
  description: string;
}
declare var $;
export default class DisplayBannerWebPart extends BaseClientSideWebPart<IDisplayBannerWebPartProps> {

  userflag: boolean = false;
  public render(): void {
    var _this = this;
    console.log(this.context.pageContext.user.loginName);
    //Checking user details in group
    checkUserinGroup("Banners", this.context.pageContext.user.email, function (result) {
     
      console.log(this.context.pageContext.user);
      if (result == 1) {
        _this.userflag = true;
      }
      _this.viewlistitemdesign();
    })
  }
  public viewlistitemdesign() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = '<section class="banner-section">' +
      '<div id="carousel-banner" class="carousel carousel-fade" data-ride="carousel">' +
      '<ol class="carousel-indicators banner-carousel">' +
      '</ol>' +
      '</div>' +
      '<div id="addEvents" class="event-add" style="Display:none">' +
      '<h3 class="banner-itemview" style="cursor:pointer">VIEW COVERAGE EVENTS <a href="' + siteURL + '/Pages/AddListItem.aspx?CName=Banners"><i class="icon-add"></i></a> </h3>' +
      '</div>' +
      '<section>';
    this.BannerPage(this.userflag);
    let viewevent = document.getElementsByClassName('banner-itemview');
    for (let i = 0; i < viewevent.length; i++) {
      viewevent[i].addEventListener("click", (e: Event) => this.viewpageRedirect(siteURL));
    }
    
  }
  viewpageRedirect(siteURL){
    window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=Banners";
  }
  BannerPage(userflag) {

    var renderhtml = '<div class="carousel-inner" role="listbox">';
    var renderliitems = "";
    var count;
    var activeflag;

    /* if (typeof this.properties.Count === "undefined") {
          count = 5;
      } else {
          count = +(this.properties.Count);
  
      }*/
    let objResults;
   var siteURL = this.context.pageContext.web.absoluteUrl;
    objResults = readItems("Banners", ["Title", "Modified", "LinkURL", "Display", "BannerContent", "Image"], 3, "Modified", "Display", 1);
    objResults.then((items: any[]) => {
      if(items.length>0)
      {
      for (let i = 0; i < items.length; i++) { 
        if (i == 0) {
          activeflag = "active";
        } else {
          activeflag = "";
        }
        renderliitems += '<li data-slide-to="' + i + '" data-target="#carousel-banner" class="' + activeflag + '">' + '</li>';
        // var reg1 = new RegExp('<div class=\"ExternalClass[0-9A-F]+\">', "");
        // var reg2 = new RegExp('</div>$', "");
        // var bancont = items[i].BannerContent.replace(reg1, "").replace(reg2, "");
        renderhtml += '<div class="item ' + activeflag + '">';
        renderhtml += '<img src="' + items[i].Image.Url + '" style="max-height: 319px;"alt="Slide" />';
        renderhtml += '<div class="carousel-caption">';
       // renderhtml += '<p> ' + bancont + '</p>';
        // if(bancont.length>65)
        // {
        //   bancont=bancont.substring(0,65)+"...";
        // }
var DottedTitle=items[i].Title;
        if(DottedTitle.length>65)
        {
          DottedTitle=DottedTitle.substring(0,65)+"...";
        }
        renderhtml += '<h3 class="wow fadeInRight" style="visibility: visible; animation-name: fadeInRight;"> ' + DottedTitle + '</h3>';
        if (items[i].LinkURL !== null) {
          renderhtml += '<div align="center">' + '<a href="' + items[i].LinkURL.Url + '" class="wow fadeInRight" style="visibility: visible; animation-name: fadeInRight;">lEARN mORE</a>' + '</div>';
        }
        renderhtml += '</div>';
        renderhtml += '</div>';
      }
    }
      else if(items.length==0){
        activeflag = "active";
        renderliitems += '<li data-slide-to="1" data-target="#carousel-banner" class="' + activeflag + '"></li>';
        renderhtml += '<div class="item ' + activeflag + '">';
        renderhtml += '<img src="' + siteURL + '/_catalogs/masterpage/BloomHomepage/images/logo.png" style="max-height: 319px;"alt="Slide" title="Slide" />';
        renderhtml += '<div class="carousel-caption">';
        renderhtml += '<p></p>';
        renderhtml += '<h3 class="wow fadeInRight no-data" style="visibility: visible; animation-name: fadeInRight;">No Banner Image To Display</h3>';
        renderhtml += '</div>';
        renderhtml += '</div>';

      }
      renderhtml += '</div>';
      ​renderhtml += '<!-- Left and right controls -->';
      renderhtml += '<a class="left carousel-control" href="#carousel-banner" data-slide="prev">';
      renderhtml += '<span class="glyphicon glyphicon-chevron-left"></span>';
      renderhtml += '<span class="sr-only">Previous</span>';
      renderhtml += '</a>';
      renderhtml += '<a class="right carousel-control" href="#carousel-banner" data-slide="next">';
      renderhtml += '<span class="glyphicon glyphicon-chevron-right"></span>';
      renderhtml += '<span class="sr-only">Next</span>';
      renderhtml += '</a>';
      //renderhtml += '<div id="addEvents" class="event-add">'+'<h3>UPDATE COVERAGE EVENTS <a href="#"><i class="icon-add"></i></a> </h3>'+'</div>';
      if (userflag == false) {
        $('#addEvents').hide();
      }
      else {
        $('#addEvents').show();
      }
      $(".banner-carousel").append(renderliitems);
      $(".banner-carousel").after(renderhtml);
      $('#carousel-banner').carousel({ interval: 8000 });
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
