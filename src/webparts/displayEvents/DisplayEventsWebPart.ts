import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DisplayEventsWebPart.module.scss';
import * as strings from 'DisplayEventsWebPartStrings';
import {
  readItems, formatDate, checkUserinGroup
} from '../../commonService';
import 'fullcalendar';
//import '../../ExternalRef/css/fullcalendar.min.css';
declare var $;

export interface IDisplayEventsWebPartProps {
  description: string;
}

export default class DisplayEventsWebPart extends BaseClientSideWebPart<IDisplayEventsWebPartProps> {

  userflag: boolean = false;

  public render(): void {
    // SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.9.0/fullcalendar.js");
    // SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.9.0/fullcalendar.css");


    var _this = this;

    //Checking user details in group
    checkUserinGroup("Events", this.context.pageContext.user.email, function (result) {
      //console.log(result);
      if (result == 1) {
        _this.userflag = true;
        _this.EventsDisplay();
        // _this.loadAllEvents();
         $('#addicon').show();
      }
      else {
        _this.EventsDisplay();
        // _this.loadAllEvents();
        $('#addicon').hide();
      }

    })
  }

  EventsDisplay() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML =
      '<div class="head-div">' +
      '<h3>Events & Holidays<a href="' + siteURL + '/Pages/ListView.aspx?CName=Events">More</a></h3>' +
      '</div>' +
      '<section class="date-section">' +
      '<div id="calendar">' +
      '<a href="' + siteURL + '/Pages/AddListItem.aspx?CName=Events" class="add-icon" id="addicon"  style="Display:none"><i class="icon-add"></i></a>' +
      '</div>' +
      '<div id="event-setting" class="carousel carousel-fade" data-ride="carousel">' +
      '<ol id="appendevents" class="carousel-indicators">' +

      '</ol>' +
      '</div>' +
      '</section>';
    // this.getHolidayEvents(this.userflag);
    document.title = "Home";
      let _that = this;
      if ($('#calendar').length > 0) {
      //setTimeout(function () {
        $('#calendar').fullCalendar({


          //   dayRender: function (date, cell) {
          //     var today = new Date();
          //     if (new Date(date).getDate() === today.getDate()) {
          //         cell.css("background-color", "red");
          //     }
          // },
  
          dayClick: function (date, jsEvent, view) {
            $('#bindEvents').remove();
            $('.carousel-dot').remove();
            _that.dateClick(_that.formatsearchdate(date));
          },
          defaultDate: new Date(),
          editable: true,
          eventLimit: true,
          contentHeight: 'auto',
          aspectRatio: 2,
          header: {
            left: 'prev',
            center: 'title',
            right: 'next'
          }
        });
      //}, 4000);
      


      $('body').on('click', 'button.fc-prev-button', function () {
        _that.getItem(new Date());
      });

      $('body').on('click', 'button.fc-next-button', function () {
        _that.getItem(new Date());
      });

    }

    _that.loadAllEvents();

  }
  arrAllDates = [];

  loadAllEvents() {
    let _that = this;
    let objResults = readItems("Events", ["Title", "Modified", "StartDate", "EndDate", "Display", "Explanation", "ExplanationText"], 5000, "Modified", "Display", 1);
    objResults.then((items: any[]) => {
      if (items && items.length) {
        items.filter(c => {
          if (c.EndDate == null) {
            c.EndDate = c.StartDate;
          }
          this.arrAllDates.push({
            Type: 'Events',
            StartDate: c.StartDate,
            EndDate: c.EndDate,
            Title: c.Title,
            Explanation: c.Explanation,
            ExplanationText: c.ExplanationText
          });
        });
      }
      _that.loadAllHolidays();
      $('#event-setting').carousel({ interval: 6000 });
    });
  }

  loadAllHolidays() {
    let _that = this;
    let objResults = readItems("Holiday", ["Title", "Modified", "EventDate", "EndEventDate", "Display"], 5000, "Modified", "Display", 1);
    objResults.then((items: any[]) => {
      if (items && items.length) {
        items.filter(c => {
          if (c.EndEventDate == null) {
            c.EndEventDate = c.EventDate;
          }
          this.arrAllDates.push({
            Type: 'Holiday',
            StartDate: c.EventDate,
            EndDate: c.EndEventDate,
            Title: c.Title
          });
        });
      }
      _that.getItem(_that.formatsearchdate(new Date()));
    });
  }

  dateClick(date) {
    var particulareventsrenderhtml = '<div class="carousel-inner"  role="listbox" id="bindEvents">';
    var particulareventsrenderliitems = "";
    var activeflag;
    var events = this.arrAllDates.filter(c => {
      var sdate = c.StartDate.split('T')[0];
      var edate;
      if (c.EndDate) {
        edate = c.EndDate.split('T')[0];
      }
      else {
        edate = sdate;
      }
      sdate = new Date(new Date(sdate).setDate(new Date(sdate).getDate() + 1))
      edate = new Date(new Date(edate).setDate(new Date(edate).getDate() + 1))
      var sdatemonth = (new Date(sdate).getMonth() + 1) + '';
      var sdateday = (new Date(sdate).getDate()) + '';
      if (sdatemonth.length == 1) {
        sdatemonth = '0' + sdatemonth;
      }
      if (sdateday.length == 1) {
        sdateday = '0' + sdateday;
      }
      var edatemonth = (new Date(edate).getMonth() + 1) + '';
      var edateday = (new Date(edate).getDate()) + '';
      if (edatemonth.length == 1) {
        edatemonth = '0' + edatemonth;
      }
      if (edateday.length == 1) {
        edateday = '0' + edateday;
      }
      var checkdatemonth = (new Date(date).getMonth() + 1) + '';
      var checkdateday = (new Date(date).getDate()) + '';
      if (checkdatemonth.length == 1) {
        checkdatemonth = '0' + checkdatemonth;
      }
      if (checkdateday.length == 1) {
        checkdateday = '0' + checkdateday;
      }

      var serverstartdate = new Date(new Date(sdate).getFullYear() + '-' + sdatemonth + '-' + sdateday);
      var serverendate = new Date(new Date(edate).getFullYear() + '-' + edatemonth + '-' + edateday);
      var checkdate = new Date(new Date(date).getFullYear() + '-' + checkdatemonth + '-' + checkdateday);

      if (c.StartDate && !c.EndDate) {
        if (serverstartdate == checkdate) {
          return c;
        }
      }
      else if (c.StartDate && c.EndDate) {
        if (serverstartdate <= checkdate && serverendate >= checkdate) {
          return c;
        }
      }
    });
    for (let i = 0; i < events.length; i++) {
      if (i == 3) {
        break;
      }
      if (i == 0) {
        activeflag = "active";
      }
      else {
        activeflag = "";
      }
      particulareventsrenderliitems += '<li data-slide-to="' + i + '" data-target="#event-setting" class="' + activeflag + '"></li>';
      var regg1 = new RegExp('<div class=\"ExternalClass[0-9A-F]+\">', "");
      var regg2 = new RegExp('</div>$', "");

      var explanation = '';
      if (events[i].Type == 'Events') {
        //explanation = events[i].ExplanationText.replace(regg1, "").replace(regg2, "");
        explanation = events[i].ExplanationText;
      }
      else if (events[i].Type == 'Holiday') {
        explanation = "&nbsp";
      }
      particulareventsrenderhtml += '<div class="item ' + activeflag + '">';
      particulareventsrenderhtml += '<div class="holiday-set">';
      if (explanation && explanation.length > 35 && events[i].Type == 'Events') {
        explanation = explanation.substring(0, 35) + "...";
      }

      if(!explanation){
        explanation = 'No Description to Display';
      }

      particulareventsrenderhtml += '<p><i class="icon-fork"></i>' + events[i].Title + '</p>';
      particulareventsrenderhtml += '</div>';
      particulareventsrenderhtml += '</div>';
    }
    if (events && !events.length) {
      particulareventsrenderhtml = '<h4 id="no-data" class="no-data" style="padding: 20px;">No Events/Holidays to display</h4>';
    }
    particulareventsrenderhtml += '</div>';
    if (this.userflag == false) {
      $('#addicon').hide();
    }
    else {
      $('#addicon').show();
    }
    $("#no-data").remove();
    $("#appendevents").empty();
    $("#appendevents").append(particulareventsrenderliitems);
    $('#bindEvents').empty();
    $("#appendevents").after(particulareventsrenderhtml);
  }


  getItem(date) {
    //$("#no-data").remove();
    if (this.arrAllDates) {
      var that = this;
      this.arrAllDates.filter(c => {
        var sdate = c.StartDate.split('T')[0];
        if (c.EndDate) {
          var edate = c.EndDate.split('T')[0];
        }
        sdate = new Date(new Date(sdate).setDate(new Date(sdate).getDate() + 1));
        edate = new Date(new Date(edate).setDate(new Date(edate).getDate() + 1));

        var sdatemonth = (new Date(sdate).getMonth() + 1) + '';
        var sdateday = (new Date(sdate).getDate()) + '';
        if (sdatemonth.length == 1) {
          sdatemonth = '0' + sdatemonth;
        }
        if (sdateday.length == 1) {
          sdateday = '0' + sdateday;
        }
        var edatemonth = (new Date(edate).getMonth() + 1) + '';
        var edateday = (new Date(edate).getDate()) + '';
        if (edatemonth.length == 1) {
          edatemonth = '0' + edatemonth;
        }
        if (edateday.length == 1) {
          edateday = '0' + edateday;
        }
        var serverstartdate = new Date(new Date(sdate).getFullYear() + '-' + sdatemonth + '-' + sdateday);
        var serverendate = new Date(new Date(edate).getFullYear() + '-' + edatemonth + '-' + edateday);

        // var servedate1:any=this.formatsearchdate(sdate);
        // var checkdate = new Date(new Date(date).getFullYear() + '-' + (new Date(date).getMonth() + 1) + '-' + new Date(date).getDate());

        if (c.StartDate && c.EndDate) {
          var totaldays = Math.round((<any>serverendate - <any>serverstartdate) / (1000 * 60 * 60 * 24));
          for (let index = 0; index <= totaldays; index++) {
            var tempstart = serverstartdate;
            var year = tempstart.getFullYear();
            var month = (tempstart.getMonth() + 1) + '';
            var day = (tempstart.getDate()) + '';
            if (month.length == 1) {
              month = '0' + month;
            }
            if (day.length == 1) {
              day = '0' + day;
            }
            var strcheckdate = year + '-' + month + '-' + day;
            $('.fc-day-grid').find('td').filter(function () {
              if ($(this).attr('data-date') == strcheckdate) {

                if ($(this).find('span').length) {
                  $(this).find('span').addClass("adddot");

                }
              }
            });
            tempstart = new Date(tempstart.setDate(tempstart.getDate() + 1));
          }
        }

        else if (c.StartDate) {
          var syear = serverstartdate.getFullYear();
          var smonth = (serverstartdate.getMonth() + 1) + '';
          var sday = (serverstartdate.getDate()) + '';
          if (smonth.length == 1) {
            smonth = '0' + smonth;
          }
          if (sday.length == 1) {
            sday = '0' + sday;
          }
          var serverstartdate2 = syear + '-' + smonth + '-' + sday;
          $('.fc-day-grid').find('td').filter(function () {
            if ($(this).attr('data-date') == serverstartdate2) {
              if ($(this).find('span').length) {
                $(this).find('span').addClass("adddot");
              }
            }
          });
        }
        this.dateClick(date);
      });
    }
  }

  formatsearchdate(date) {
    var ddate = new Date(date);
    var yr = ddate.getFullYear();
    var month = ddate.getMonth() + 1;
    var day = ddate.getDate() < 10 ? '0' + ddate.getDate() : ddate.getDate();
    var newDate = (month < 10 ? '0' + month.toString() : month) + '/' + day + '/' + yr;
    return newDate;
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
