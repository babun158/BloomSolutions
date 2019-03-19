import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DisplayAnnouncemnetWebPart.module.scss';
import * as strings from 'DisplayAnnouncemnetWebPartStrings';
import {
  readItems, formatDate
} from '../../commonService';
import pnp from "sp-pnp-js";
declare var $;
export interface IDisplayAnnouncemnetWebPartProps {
  description: string;
}

export default class DisplayAnnouncemnetWebPart extends BaseClientSideWebPart<IDisplayAnnouncemnetWebPartProps> {

  public render(): void {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML =
      '<section class="announce">' +
      '<div class="head-div ">' +
      '<h3>' +
      '<i class="icon-list"></i>' + "Announcements" + '<a href="' + siteURL + '/Pages/ListView.aspx?CName=Announcements">More</a>' +
      '</h3>' +
      '</div>' +
      '<div class="annouce-bg" id="bannerimageId">' +
      '<div class="announce-carousel">' +
      '</div>' +
      '</div>' +
      '</section>';

    this.GetAnnouncementsNews();
  }
  GetAnnouncementsNews() {

    var announcementsrenderhtml = '<div id="bindclickdata" class="annouce-list col-lg-6 col-md-6 col-sm-6">';
    var count;
    var setid = 0;
    var activeflag;
    let objectResults;
    var siteURL = this.context.pageContext.web.absoluteUrl;
    objectResults = readItems("Announcements", ["ID", "Title", "Modified", "Explanation","ExplanationText", "Image", "Expires", "ViewedUsers"], 3, "Modified", "Display", 1);
    objectResults.then((items: any[]) => {
      if (items && items.length > 0) {
        for (let i = 0; i < items.length; i++) {
          setid = items[i].ID;
          if (i >= 0) {
            if (i == 0) {
              announcementsrenderhtml += '<p>' + formatDate(items[i].Modified) + '</p>';
              announcementsrenderhtml += '<a href="' + siteURL + '/Pages/Viewlistitem.aspx?CName=Announcements&CID=' + items[i].ID + '" class="head5">' + items[i].Title + '</a>';
              if (items[i].ExplanationText.length > 35) {
                items[i].ExplanationText = items[i].ExplanationText.substring(0, 35) + "...";
              }
              announcementsrenderhtml += '<p>' + items[i].ExplanationText + '</p>';
              announcementsrenderhtml += '<div align="center"><ul><li><i class="icon-eye views' + items[i].ID + '"></i></li><li><i class="icon-comments cmd' + items[i].ID + '"></i></li><li><i class="icon-heart likes' + items[i].ID + '"></i></li></ul></div></div>';
              announcementsrenderhtml += '<div class="event-list col-lg-6 col-md-6 col-sm-6"><ul>';
              //$('#bannerimageId').css('background-image', 'url(' + items[i].Image.Url + ')');
              $('#bannerimageId').css({'background-image': 'url(' + items[i].Image.Url + ')','background-size': 'cover'});
              if (items[i].ExplanationText.length > 20) {
                items[i].ExplanationText = items[i].ExplanationText.substring(0, 20) + "...";
              }
              announcementsrenderhtml += '<li class="newschange" id=' + items[i].ID + '>' + '<a id=' + items[i].ID + ' href="#"><img src=' + items[i].Image.Url + ' />' + items[i].ExplanationText + ' <span>' + formatDate(items[i].Modified) + '</span></a>' + '</li>';
            }
            else {
              if (items[i].ExplanationText.length > 20) {
                items[i].ExplanationText = items[i].ExplanationText.substring(0, 20) + "...";
              }
              announcementsrenderhtml += '<li class="newschange" id=' + items[i].ID + '>' + '<a id=' + items[i].ID + ' href="#"><img src=' + items[i].Image.Url + ' />' + items[i].ExplanationText + ' <span>' + formatDate(items[i].Modified) + '</span></a>' + '</li>';
            }
          }
        }
        announcementsrenderhtml += '</ul>';
        announcementsrenderhtml += '</div>';
      }
      else {
        announcementsrenderhtml += '<h3 class="no-data" style="color:white;">No Announcements To Display</h3>';
        announcementsrenderhtml += '</div>';
      }
      $(".announce-carousel").append(announcementsrenderhtml);
      var ViewedUsers = 0;
      for (let i = 0; i < items.length; i++) {
        if (items[i].ViewedUsers &&items[i].ViewedUsers.split(',').length >0)
        {
          ViewedUsers = items[i].ViewedUsers.split(',').length;
        }  
        $('.views' + items[i].ID).after("<a>" + ViewedUsers + "</a>");
        var commentscount = 0;
        var objResults1 = readItems("AnnouncementComments", ["AnnouncementID"], 1000, "Modified", "AnnouncementID", items[i].ID);
        objResults1.then((itemsCount: any[]) => {
          if (itemsCount && itemsCount.length > 0) {
            commentscount = itemsCount.length;
          }
          $('.cmd' + items[i].ID).after("<a>" + commentscount + "</a>");
        });
        //var Likescount = 0;
        var objResults2 = readItems("AnnouncementsLikes", ["AnnouncementID","Liked"], 1000, "Modified", "AnnouncementID", items[i].ID);
        objResults2.then((itemsCount2: any[]) => {
          let LikesCount=0;
          if (itemsCount2 && itemsCount2.length > 0) {
          for(let j=0;j<itemsCount2.length;j++){
          if(itemsCount2[j].Liked==true){
          LikesCount++;
          }
          }
        }
          // if (itemsCount2 && itemsCount2.length > 0) {
          //   LikesCount = itemsCount2.length;
          // }
          $('.likes' + items[i].ID).after("<a>" + LikesCount + "</a>");
        });

      }


      let Addevent = document.getElementsByClassName('newschange');
      for (let i = 0; i < Addevent.length; i++) {
        Addevent[i].addEventListener("click", (e: Event) => this.bindtodiv(Addevent[i].id));
      }
    });
  }
  bindtodiv(changeID) {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    var bindannouncementsrenderhtml = "";
  // bindannouncementsrenderhtml = '<div id="bindclickdata" class="annouce-list col-lg-6 col-md-6 col-sm-6">';
    pnp.sp.web.lists.getByTitle("Announcements").items.getById(changeID).get()
      .then((results: any) => {
        bindannouncementsrenderhtml += '<p>' + formatDate(results.Modified) + '</p>';
        bindannouncementsrenderhtml += '<a href="' + siteURL + '/Pages/Viewlistitem.aspx?CName=Announcements&CID=' + changeID + '" class="head5">' + results.Title + '</a>';
        if (results.ExplanationText.length > 35) {
          results.ExplanationText = results.ExplanationText.substring(0, 35) + "...";
        }
        bindannouncementsrenderhtml += '<p>' + results.ExplanationText + '</p>';

        bindannouncementsrenderhtml += '<div align="center"><ul><li><i class="icon-eye views' + changeID + '"></i></li><li><i class="icon-comments cmd' + changeID + '"></i></li><li><i class="icon-heart likes' + changeID + '"></i></li></ul></div>';
       // bindannouncementsrenderhtml += '</div>';
        //$('#bannerimageId').css('background-image', 'url(' + results.Image.Url + ')');
        $('#bannerimageId').css({'background-image': 'url(' + results.Image.Url + ')','background-size': 'cover'});       
        $("#bindclickdata").empty();
        $("#bindclickdata").append(bindannouncementsrenderhtml);
        let ViewedUsers = 0;
        if (results.ViewedUsers && results.ViewedUsers.split(',').length > 0) {
          ViewedUsers = results.ViewedUsers.split(',').length;
        }
        $('.icon-eye').empty();
        $('.views' + changeID).append("<a>" + ViewedUsers + "</a>");
        //var itemsCount = 0;
        var commentscount = 0;
        var objResults1 = readItems("AnnouncementComments", ["AnnouncementID"], 1000, "Modified", "AnnouncementID", changeID);
        objResults1.then((itemsCount: any[]) => {
         
          if (itemsCount && itemsCount.length > 0) {
            commentscount = itemsCount.length;
          }
          $('.icon-comments').empty();
          $('.cmd' + changeID).append("<a>" + commentscount + "</a>");
        });
       // var Likescount = 0;
        var objResults2 = readItems("AnnouncementsLikes", ["AnnouncementID","Liked"], 1000, "Modified", "AnnouncementID", changeID);
        objResults2.then((itemsCount2: any[]) => {
          let LikesCount=0;
          if (itemsCount2 && itemsCount2.length > 0) {
          for(let j=0;j<itemsCount2.length;j++){
          if(itemsCount2[j].Liked==true){
          LikesCount++;
          }
          }
        }
          // if (itemsCount2 && itemsCount2.length > 0) {
          //   LikesCount = itemsCount2.length;
          // }
          $('.icon-heart').empty();
          $('.likes' + changeID).append("<a>" + LikesCount + "</a>");
        });
        
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
