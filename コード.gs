var ADMIN_SPREADSHEET_ID = '1evJKaau2XenllewyQ6BJnXT6WqS1wE5WSYFbn4O3eko';
var ADMIN_SHEET_NAME = 'User';
var ACTION_URL = 'https://script.google.com/macros/s/AKfycbx-j7R1EuJOlqN2pJ47Y0TKCwTaATLhKuSUgSWKs-p5B7FuRYE/exec';

function Work(param){
  this.category = param.category;
  this.name = param.name;
  this.point = param.point;
  this.description = param.description;
  this.lastdate = param.lastdate;
  this.row = param.row;
}

Work.get_work_list = function(spreadsheet){
  var work_list = []; 
  var sheet = spreadsheet.getSheetByName('work');
  var values = sheet.getDataRange().getDisplayValues();
  for(var index = 1; index < values.length; index++){
    var work = new Work({
      category: values[index][0],
      name: values[index][1],
      point: values[index][2],
      description: values[index][3],
      row: index + 1
    });
    work_list.push(work);
  }
  
  function compare(a, b) {
  // Use toUpperCase() to ignore character casing
    var categoryA = a.category.toUpperCase();
    var categoryB = b.category.toUpperCase();
  
    var comparison = 0;
    if (categoryA > categoryB) {
      comparison = 1;
    } else if (categoryA < categoryB) {
      comparison = -1;
    }
    return comparison;
  }
  work_list.sort(compare);
  Logger.log("work_list");
  //Logger.log(work_list);
  return work_list;
};

Work.prototype.set = function(spreadsheet) {
  var sheet = spreadsheet.getSheetByName('work');
  var values = [this.category, this.name, this.point, this.description];
  if(this.category != undefined){
    sheet.appendRow(values);
  }
  Logger.log("set work");
};

//サイトアクセス時にHTMLページを渡す
function doGet(e) {
  var html_output;
  var user_email = Session.getActiveUser().getUserLoginId();
  var spreadsheet_url = get_spreadsheet_url(user_email);
  
  //authorize user
  try{
    var spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
  }catch(error){
    var message = error.message;
  }finally{
    if(spreadsheet == undefined){
      html_output = get_setting_output(spreadsheet_url, message);
      return html_output;
    }
  }

  switch(e.parameter.path){
    case 'history':
      html_output = get_history_output(spreadsheet);
      break;
    case 'setting':
      html_output = get_setting_output(spreadsheet_url);
      break;
    case 'work':
      html_output = get_work_output(spreadsheet);
      break;
    default:
      html_output = get_index_output(spreadsheet, user_email);
      break;
  }
  return html_output;
}

//データPOST時にhistoryに書き込む
function doPost(e) {
  var html_output;
  Logger.log(e.parameter.spreadsheet_url);
  Logger.log("Recieved POST. data-type is ");
  Logger.log(e.parameter.data_type);
  switch(e.parameter.data_type){
    case 'setting':
      var is_registed = false;
      var user_email = Session.getActiveUser().getUserLoginId();
      var spreadsheet_admin = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
      var sheet_admin = spreadsheet_admin.getSheetByName(ADMIN_SHEET_NAME);
      var spreadsheet_url = e.parameter.spreadsheet_url;
      var last_row = sheet_admin.getLastRow();
      var last_column = sheet_admin.getLastColumn(); 
      for(var row = 1; row <= last_row; row++){
        var email = sheet_admin.getRange(row, 1).getDisplayValue();
        if(email == user_email){
          is_registed = true;
          sheet_admin.getRange(row, 2).setValue(spreadsheet_url);
        }
      }
      if(is_registed == false){
        sheet_admin.getRange(last_column+1, 1).setValue(user_email);
        sheet_admin.getRange(last_column+1, 2).setValue(spreadsheet_url);
      }
      try{
        var spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
      }catch(error){
        var message = error.message;
      }finally{
        if(spreadsheet == undefined){
          html_output = get_setting_output(spreadsheet_url, message);
        }else{
          html_output = get_index_output(spreadsheet, user_email);
        }
      }
      break;
    case 'submit_work':
    Logger.log(e);
      var spreadsheet_url = e.parameter.spreadsheet_url;
      var spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
      var recode = {
        'date': e.parameter.date,
        'category': e.parameter.category,
        'work': e.parameter.work,
        'user': e.parameters.user.join(),
        'point': e.parameter.point
      };
      set_history(spreadsheet, recode);
      html_output = get_history_output(spreadsheet);
      break;
    case 'delete':
      var spreadsheet_url = e.parameter.spreadsheet_url;
      var spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
      delete_history(spreadsheet, e.parameter.index);
      html_output = get_history_output(spreadsheet);
    case 'add_work':
      var user_email = Session.getActiveUser().getUserLoginId();
      var spreadsheet_url = e.parameter.spreadsheet_url;
      var spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
      var work = new Work({
        category: category,
        name: e.parameter.name,
        point: e.parameter.point,
        description: e.parameter.description
      });
      work.set(spreadsheet);
      html_output = get_index_output(spreadsheet, user_email);
    default:
      break;
  }
  return html_output;
}

function auto_submit() {
  var spreadsheet_admin = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  var sheet_admin = spreadsheet_admin.getSheetByName(ADMIN_SHEET_NAME);
  var values_admin = sheet_admin.getDataRange().getValues();
  var date = new Date();
  var date_today = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
  var done_url = [];
  for(var index = 0; index < values_admin.length; index++){
    var spreadsheet_url = values_admin[index][1];
    if(done_url.indexOf(spreadsheet_url) == -1){
      var spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
      var sheet = spreadsheet.getSheetByName('setting');
      var values = sheet.getRange(5, 2, 4, sheet.getRange(5, sheet.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn()).getValues();
      Logger.log(values);
      for(var col = 0; col < values[0].length - 1; col++){
        var recode = {
          'date': date_today,
          'category': values[0][col],
          'work': values[1][col],
          'user': values[2][col],
          'point': values[3][col]
        }
        Logger.log(recode);
        set_history(spreadsheet, recode);
      }
      done_url.push(spreadsheet_url);
    }
  }
}

function set_work(spreadsheet, category, work, point, description){
  var sheet = spreadsheet.getSheetByName('work');
  var values = [category, work, point, description];
  if(category != undefined){
    sheet.appendRow(values);
  }
  Logger.log("set work");
}

function set_history(spreadsheet, recode){
  var sheet = spreadsheet.getSheetByName('history');
  var values = [
    recode["date"],
    recode["category"],
    recode["work"],
    recode["user"],
    recode["point"]
  ];
  sheet.appendRow(values);
}

function delete_history(spreadsheet, index){
  var sheet = spreadsheet.getSheetByName('history');
  sheet.deleteRow(parseInt(index) + 1);
  Logger.log("delete history");
  Logger.log(index + 1);
}

function get_index_output(spreadsheet, user_email){
  var tpl = HtmlService.createTemplateFromFile('index');
  var group_dict = get_group_dict(spreadsheet);
  var history_list = get_history_list(spreadsheet);
  tpl.group_dict = group_dict;
  tpl.total_dict = get_total_dict(group_dict, history_list);
  tpl.work_list = Work.get_work_list(spreadsheet);
  tpl.action_url = ACTION_URL;
  tpl.user_email = user_email;
  var output = tpl.evaluate();
  output = output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

function get_history_output(spreadsheet){
  var tpl = HtmlService.createTemplateFromFile('history');
  var group_dict = get_group_dict(spreadsheet);
  var history_list = get_history_list(spreadsheet);
  tpl.group_dict = group_dict;
  tpl.history_list = history_list;
  tpl.total_dict = get_total_dict(group_dict, history_list);
  tpl.action_url = ACTION_URL;
  var output = tpl.evaluate();
  output = output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

function get_setting_output(spreadsheet_url, message){
  var tpl = HtmlService.createTemplateFromFile('setting');
  tpl.spreadsheet_url = spreadsheet_url;
  tpl.action_url = ACTION_URL;
  tpl.message = message;
  var output = tpl.evaluate();
  output = output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

function get_work_output(spreadsheet){
  var tpl = HtmlService.createTemplateFromFile('work');
  var group_dict = get_group_dict(spreadsheet);
  tpl.group_dict = group_dict;
  tpl.action_url = ACTION_URL;
  tpl.work_list = Work.get_work_list(spreadsheet);
  var output = tpl.evaluate();
  output = output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

function get_error_output(message){
  var tpl = HtmlService.createTemplateFromFile('error');
  tpl.message = message;
  var output = tpl.evaluate();
  output = output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

function get_spreadsheet_url(user_email) {
  var spreadsheet_admin = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  var sheet_admin = spreadsheet_admin.getSheetByName(ADMIN_SHEET_NAME);
  var values = sheet_admin.getDataRange().getValues();
  for(var index = 0; index < values.length; index++){
    var email = values[index][0];
    if(email == user_email){
      var spreadsheet_url = values[index][1];
      return spreadsheet_url;
    }
  }
  return undefined;
}

function get_group_dict(spreadsheet) {
  var member_dict = {};
  var group_dict= {};
  var spreadsheet_url = spreadsheet.getUrl();
  var sheet = spreadsheet.getSheetByName('setting');
  var values = sheet.getDataRange().getValues();
  var group_name = values[1][1];
  for(var index = 1; index < values[3].length; index++){
    var email = values[2][index];
    var name = values[3][index];
    member_dict[email] = {
      "name": name,
      "group_name": group_name,
      "spreadsheet_url": spreadsheet_url
    }
  }
  group_dict['group_name'] = group_name;
  group_dict['spreadsheet_url'] = spreadsheet_url;
  group_dict['member'] = {};
  for(email in member_dict){
    group_dict['member'][email] = member_dict[email];
  }
  Logger.log("group_dict");
  return group_dict;
}

function get_history_list(spreadsheet){
  var history_list = []; 
  var sheet = spreadsheet.getSheetByName('history');
  var values = sheet.getDataRange().getDisplayValues();
  for(var row = values.length - 1; row > 0; row--){
    var date = values[row][0];
    var category = values[row][1];
    var work = values[row][2];
    var name = values[row][3];
    var point = values[row][4];
    history_list.push({
      'index': row,
      'date': date,
      'category': category,
      'work': work,
      'name': name,
      'point': point
    });
  }
  Logger.log("history_list");
  Logger.log(history_list);
  return history_list;
}

function get_work_dict(spreadsheet){
  var work_dict = {}; 
  var sheet = spreadsheet.getSheetByName('work');
  var values = sheet.getDataRange().getDisplayValues();
  for(var index = 1; index < values.length; index++){
    var category = values[index][0];
    var name = values[index][1];
    var point = values[index][2];
    var description = values[index][3];
    if(!(category in work_dict)){
      work_dict[category] = {};
    }
    work_dict[category][name] = {'point': point,'description': description};
  }
  Logger.log("work_dict");
  return work_dict;
}

function get_total_dict(group_dict, history_list){
  var total_dict = {};
  for(var email in group_dict['member']){
    total_dict[group_dict['member'][email]['name']] = 0;
  }
  var regexp = new RegExp(Object.keys(total_dict).join('|'), 'g');
  history_list.forEach(function(recode){
    var names = recode['name'].match(regexp);
    names.forEach(function(name){
      var point = parseInt(recode['point'], 10) / names.length;
      total_dict[name] += point;
    });
  });
  for(name in total_dict){
    total_dict[name] = Math.round(total_dict[name]);
  }
  Logger.log("total_dict");
  //Logger.log(total_dict);
  return total_dict;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

