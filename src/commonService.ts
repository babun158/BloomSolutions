
declare var $;
import pnp from "sp-pnp-js";
function readItems(listName: string, listColumns: string[], topCount: number, orderBy: string, filterKey?: string, filterValue?: any): Promise<any> {
    var matchColumns = formString(listColumns);
    if (filterKey == undefined) {
        return pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).top(topCount).orderBy(orderBy).get();
    }
    else {
        return pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).filter("" + filterKey + " eq '" + filterValue + "'").top(topCount).orderBy(orderBy).get();
    }
}

function addItems(listName: string, listColumns: any): Promise<any> {
    return pnp.sp.web.lists.getByTitle(listName).items.add(listColumns);
}

function updateitems(listName: string, id: number, listColumns: any): Promise<any> {
    return pnp.sp.web.lists.getByTitle(listName).items.getById(id).update(listColumns);
}
var batch;
function batchDelete(listName: string, selectedArray: number[], callback) {
    batch = pnp.sp.createBatch();
    for (var i = 0; i < selectedArray.length; i++) {
        pnp.sp.web.lists.getByTitle(listName).items.getById(selectedArray[i]).inBatch(batch).delete().then(r => {
            //console.log(r);
        });
    }
    callback(batch);
}

function formString(listColumns: string[]) {
    var formattedString: string = "";
    for (let i = 0; i < listColumns.length; i++) {
        formattedString += listColumns[i] + ',';
    }
    return formattedString.slice(0, -1);
}

function formatDate(dateVal) {
    var date = new Date(dateVal);
    var year = date.getFullYear();
    var locale = "en-us";
    var month = date.toLocaleString(locale, { month: "long" });
    var dt = date.getDate();
    var dateString: string;
    if (dt < 10) {
        dateString = "0" + dt;
    }
    else
        dateString = dt.toString();

    return dateString + ' ' + month.substr(0, 3) + ' ' + year;
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

export { readItems, addItems, updateitems, batchDelete, formString, formatDate, checkUserinGroup };