import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DisplayPollingsWebPart.module.scss';
import * as strings from 'DisplayPollingsWebPartStrings';
//import * as CanvasJS from "canvasjs"
//import { GoogleCharts } from 'google-charts';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { addItems, readItems, checkUserinGroup } from '../../commonJS';
import pnp from "sp-pnp-js";
export interface IDisplayPollingsWebPartProps {
  description: string;
}
declare var $;
//declare var google;
declare var alertify: any;
export default class DisplayPollingsWebPart extends BaseClientSideWebPart<IDisplayPollingsWebPartProps> {

  siteURL: string;
  userflag: boolean = false;
  public render(): void {
    this.siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = "<div class='head-div'>" +
      "<h3>vote <a href='" + this.siteURL + "/Pages/ListView.aspx?CName=Polls'>More</a></h3>" +
      "</div>" +
      "<div class='vote'>" +
      "</div>";

    this.fetchItems();

  }
  async  AddPollResult() {
    var isAllfield = true;
    let QuestionId = $('.vote').attr('ques-id');
    let curntUsrLogName = this.context.pageContext.user.loginName;
    let response = await pnp.sp.web.lists.getByTitle("PollsResults").items.select("Names", "ID").filter("Names eq '" + curntUsrLogName + "' and QuestionID eq '" + QuestionId + "'").get();
    //response.length > 0 ? $('.vote').find('a').hide() : $('.vote').find('a').show();

    var selradio = false;
    $("input[type='radio']").filter(c => {
      if (!selradio) {
        selradio = $("input[type='radio']:eq(" + c + ")").prop("checked");
      }
    });
    if (!selradio && !this.isAdmin) {
      alertify.set('notifier', 'position', 'top-right');
      alertify.error("Please select atleast one option");
      isAllfield = false;
      return;
    }
    else if (response.length > 0) {
      // alertify.alert()
      //   .setting({
      //     'label': 'OK',
      //     'message': 'You have already voted',
      //   }).show().set('closable', false).setHeader('Message');
      isAllfield = false;
      return;
    }

    let objPol = {
      Question: $('#qusdisplay').text(),
      QuestionID: QuestionId,
      Options: $('.radio input:checked').next().text(),
      Names: this.context.pageContext.user.loginName
    };
    if (isAllfield) {
      var _this = this;
      addItems('PollsResults', objPol).then(function (result) {
        _this.checkIfAlreadyVoted();
      });
      var closable = alertify.alert().setting('closable');
      //grab the dialog instance using its parameter-less constructor then set multiple settings at once.
      alertify.alert()
        .setting({
          'label': 'OK',
          'message': 'Voted SucessFully..!',
        }).show().set('closable', false).setHeader('Message');
      // alertify.confirm('Voted Successfully..!', function(){
      //    alertify.success('Ok') 
      //   });
    }

  }

  isAdmin = false;
  checkIfAlreadyVoted() {
    var _this = this;
    var userflag = false;
    let QuestionId = $('.vote').attr('ques-id');
    let curntUsrLogName = this.context.pageContext.user.loginName;
    let CName = "Polls"
    checkUserinGroup("Admin", this.context.pageContext.user.email, function (result) { //var _this = this;
      if (result == 1) {
        _this.isAdmin = true;
        $('#makeVote').text('View all results');
        $('#makeVote').attr('href', _this.siteURL + '/Pages/PollsAdminView.aspx');
      } else {
        pnp.sp.web.lists.getByTitle("PollsResults").items.select("Names", "ID").filter("Names eq '" + curntUsrLogName + "' and QuestionID eq '" + QuestionId + "'").get().then(function (result) {
          if (result && result.length > 0) {
            $('#makeVote').text('View results');
            $('#makeVote').attr('href', _this.siteURL + '/Pages/PollsView.aspx?CName=' + CName + '&CID=' + QuestionId);

          }
          else {
            $('#makeVote').text('Mark Your Vote');
            let Addevent = $('.vote').find('a');
            Addevent.on("click", (e: Event) => _this.AddPollResult());
          }
        }).catch(function () {
          $('#makeVote').text('Mark Your Vote');
          let Addevent = $('.vote').find('a');
          Addevent.on("click", (e: Event) => _this.AddPollResult());
        });
      }
    });
  }

  async fetchItems() {
    let columnArray: any = ["ID", "Question", "Options"];
    let question = await readItems("Polls", columnArray, 1, "Display", "Display", 1);
    var pollOptions = "";
    var pollMark = "";
    if (question.length > 0) {
      $('.vote').attr('ques-id', question[0].ID);
      $('.vote').append("<h5 id='qusdisplay' title='"+question[0].Question+"'>" + question[0].Question.substring(0, 50) + "</h5><div class='radio-btn'>");
      var optionsArray = question[0].Options.trim().split(';');

      var newArray = optionsArray.filter(function (v) {
        return /\S/.test(v);
      });
      $.each(newArray, function (ind, val) {
        if (val.length > 18) {
          val = val.substring(0, 18) + "...";
        }

        pollOptions += "<div class='radio'><input id='radio-" + (ind + 1) + "' name='radio' type='radio'><label for='radio-" + (ind + 1) + "' class='radio-label'>" + val + "</label></div>";
      });
      $('.radio-btn').append(pollOptions + "</div>");
      // pollMark += "<a href='#' id='makeVote'>Mark Your Vote</a>";
      pollMark += "<a href='#' id='makeVote'></a>";

      $('.radio-btn').after(pollMark);
      // let Addevent = $('.vote').find('a');
      // Addevent.on("click", (e: Event) => this.AddPollResult());

      this.checkIfAlreadyVoted();

    }
    else {
      $('.vote').append("<h5></h5><div class='radio-btn'>");
      pollOptions += "<div><h4 class='no-data on-data-polls'>No Polls to display</h4></div>";
      $('.radio-btn').append(pollOptions + "</div>");
    }

    // let QuestionId =$('.vote').attr('ques-id');
    // let curntUsrLogName = this.context.pageContext.user.loginName;
    // let response = await pnp.sp.web.lists.getByTitle("PollsResults").items.select("Name", "ID").filter("Name eq '" + curntUsrLogName + "' and QuestionID eq '" + QuestionId + "'").get();
    // //response.length > 0 ? $('.vote').find('a').hide() : $('.vote').find('a').show();
    // if(response.length>0)
    // {
    //     alertify.set('notifier', 'position', 'top-right');
    //     alertify.error("You Have Already Voted..!");
    // }
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
