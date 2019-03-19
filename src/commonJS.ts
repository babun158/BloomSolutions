import pnp from "sp-pnp-js";
declare var $;

async function addItems(listName: string, listColumns: any) {
  var resultData: any = await pnp.sp.web.lists.getByTitle(listName).items.add(listColumns);
}

async function readItems(listName: string, listColumns: string[], topCount: number, orderBy: string, filterKey?: string, filterValue?: any) {
  var matchColumns = formString(listColumns);
  var resultData: any;
  if (filterKey == undefined) {
     resultData = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).top(topCount).orderBy(orderBy).get()
  }
  else {
     resultData = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).filter("" + filterKey + " eq '" + filterValue + "'").top(topCount).orderBy(orderBy).get()
  }
  return (resultData);
}

function checkUserinGroup(Componentname: string, email: string, callback) {
  var myitems: any[];
  pnp.sp.web.siteUsers
      .getByEmail(email)
      .groups.get()
      .then((items: any[]) => {
          var currentComponent = Componentname;
          myitems = $.grep(items, function (obj, index) {
              if (obj.Title.indexOf(currentComponent) != -1) {
                  return true;
              }
          });
          callback(myitems.length);
      });
}

function formString(listColumns: string[]) {
  var formattedString: string = "";
  for (let i = 0; i < listColumns.length; i++) {
    formattedString += listColumns[i] + ',';
  }
  return formattedString.slice(0, -1);
}

export { addItems, readItems,checkUserinGroup }