/* GTM API methods */

var data = {};

function fetchAccounts() {
  var accounts = TagManager.Accounts.list({
    fields: 'account(accountId,name)'
  }).account;
  return accounts || [];
}

function fetchContainers(aid) {
  var parent = 'accounts/' + aid;
  var containers = TagManager.Accounts.Containers.list(parent, {
    fields: 'container(accountId,containerId,publicId,name)'
  }).container;
  return containers || [];
}

function fetchVersions(aid, cid) {
  var parent = 'accounts/' + aid + '/containers/' + cid;
  var versions = TagManager.Accounts.Containers.Version_headers.list(parent).containerVersionHeader;
  return versions || [];
}

function fetchVersion(aid, cid, vid) {
  var parent = 'accounts/' + aid + '/containers/' + cid + '/versions/' + vid;
  return TagManager.Accounts.Containers.Versions.get(parent);
}

function fetchTriggers(aid, cid) {
  
  var parent = 'accounts/' + aid + '/containers/' + cid;
  var workspaces = TagManager.Accounts.Containers.Workspaces.list(parent).workspace;
  
  //SpreadsheetApp.getUi().alert('Workspaces: ' + JSON.stringify(workspaces));
  var dflt = workspaces.filter(ws => ws.name == 'Default Workspace')[0];
  
  parent += '/workspaces/' + dflt.workspaceId;
  return TagManager.Accounts.Containers.Workspaces.Triggers.list(parent).trigger;
}

function createVersion(aid, cid, wsid) {
  return TagManager.Accounts.Containers.Workspaces.create_version({"name": "Created by GTM Tools Google Sheets add-on", "notes": "Created by GTM Tools Google Sheets add-on"}, 'accounts/' + aid + '/containers/' + cid + '/workspaces/' + wsid).containerVersion;
}

function getWorkspaces() {
  var apiPath = getApiPath();
  
  if (!apiPath) {
    return false;
  }
  
  Logger.log('getWorkspaces: ' + apiPath);
  return TagManager.Accounts.Containers.Workspaces.list(apiPath).workspace;
}

function fetchContainersWithSelectedMarked(aid) {
  var containerSummary = fetchContainers(aid);
  var selectedContainerId = getContainerIdFromApiPath();
  containerSummary.forEach(function(cont) {
    cont.selected = cont.containerId === selectedContainerId;
  });
  return containerSummary;
}

function fetchAccountsWithSelectedMarked() {
  var accountSummary = fetchAccounts();
  var selectedAccountId = getAccountIdFromApiPath();
  accountSummary.forEach(function(acct) {
    acct.selected = acct.accountId === selectedAccountId;
  });
  return accountSummary;
}

function getContainerPublicIdFromSheetName() {
  var sheet = SpreadsheetApp.getActiveSheet().getName();
  var cid = sheet.match(/^GTM-[a-zA-Z0-9]{4,}/) || [];
  return cid.length ? cid[0] : 'N/A';
}

function getAccountIdFromApiPath() {
  var apiPath = getApiPath();
  return apiPath ? apiPath.split('/')[1] : '';
}

function getContainerIdFromApiPath() {
  var apiPath = getApiPath();
  return apiPath ? apiPath.split('/')[3] : '';
}

function getApiPath() {
  
  var active = SpreadsheetApp.getActiveSheet().getName();
  
  if (!/^GTM-[a-zA-Z0-9]{4,}_(container|tags|variables|triggers)$/.test(active)) {
    return false;
  }
  
  var containerSheet = SpreadsheetApp.getActive().getSheetByName(active.replace(/_.+$/,'_container'));
  var apiPath = containerSheet.getRange('B6').getValue() + '';
  
  if (apiPath && apiPath.indexOf('/versions') != -1)
     apiPath = apiPath.replace(/\/versions\/.*/, '');
  
  return apiPath;
}

function getVersionCount() {
  
  var active = SpreadsheetApp.getActiveSheet().getName();
  
  if (!/^GTM-[a-zA-Z0-9]{4,}_(container|tags|variables|triggers)$/.test(active)) {
    return false;
  }

  var containerSheet = SpreadsheetApp.getActive().getSheetByName(active.replace(/_.+$/,'_container'));
  return containerSheet.getRange('A2').getValues().filter(v => v[0] == 'Version ID:').length;
}

function insertSheet(sheetName, mode) {
  
  Logger.log('insertSheet:', sheetName, mode);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var ui = SpreadsheetApp.getUi();
  var response;
  
  if (!sheet) 
    return SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  
  _clear = () => {
    sheet.clear();
    try {
        sheet.deleteRows(1, sheet.getMaxRows() - 1);
    } catch(ex) {}
  }
  
  if (mode == 'rebuild') {
    _clear();
    return sheet;
  } 
  
  response = ui.alert('Sheet named ' + sheetName + ' already exists! Click OK to overwrite, CANCEL to skip.', ui.ButtonSet.OK_CANCEL);
    
  if (response === ui.Button.OK) {
    
    _clear();
    return sheet;
  }  
    
  return false;
}

function getAssetOverview(assets) {
  var assetlist = {};
  var sortedlist = [];
  var sum = 0;  
  assets.forEach(function(item) {
    if (!assetlist[item.type]) {
      assetlist[item.type] = 1;
    } else {
      assetlist[item.type] += 1;
    }
    sum += 1;
  });
  for (var item in assetlist) {
    sortedlist.push([item, assetlist[item]]);
  }
  sortedlist = sortedlist.sort(function(a,b) {
    return b[1] - a[1];
  });
  return {
    sortedlist: sortedlist.length === 0 ? [['','']] : sortedlist,
    sum: sum
  }
}

function buildRangesObject() {
  
  //Logger.log('buildRangesObject');
  
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  var ranges = {};
  
  namedRanges.forEach(nrange => {
    
    var name = nrange.getName();
    
    if (/(_ids|_folders|_notes|_json)$/.test(name)) {
      
      var bareName = name.replace(/(_ids|_folders|_notes|_json)$/g, '');
      ranges[bareName] = ranges[bareName] || {};

      if (/_folders$/.test(name)) {
        ranges[bareName].folders = nrange.getRange();
      }
      
      if (/_ids$/.test(name)) {
        ranges[bareName].ids = nrange.getRange();
      }
      
      if (/_notes$/.test(name)) {
        ranges[bareName].notes = nrange.getRange();
      }
      
      if (/_json$/.test(name)) {    
        ranges[bareName].json = nrange.getRange();
      }
     
      ranges[bareName].type = name.split('_')[0];
      ranges[bareName].accountId = name.split('_')[1];
      ranges[bareName].containerId = name.split('_')[2];
    }
    
  });
  
  for (var i in ranges) {
  
      var folders = namedRanges.filter(nr => nr.getName() == 'folders_' + ranges[i].accountId + '_' + ranges[i].containerId);
     
      if (folders.length) {
          ranges[i].lookup_folders = folders[0].getRange();
          //Logger.log('Folder lookup added:', folders[0].getName());
      }
  };
  
  return ranges;
}

function updateSingle(item, wsid, items) {
  
  var Type = item.type[0].toUpperCase() + item.type.slice(1);
  var path = 'accounts/' + item.accountId + '/containers/' + item.containerId + '/workspaces/' + wsid + '/' + item.type + '/' + item.id;
  
  Logger.log('updateSingle getting:', item, Type, path);
  
  var remote = TagManager.Accounts.Containers.Workspaces[Type].get(path);
  
  //if (items.length > 6)
    Utilities.sleep(4.02 * 1000);
  
  if (item.notes)
    remote.notes = item.notes;
  
  if (item.fid)
    remote.parentFolderId = item.fid;
  
  Logger.log('updateSingle updating:', item, remote);
  TagManager.Accounts.Containers.Workspaces[Type].update(remote, path);
  
  //if (items.length > 6)
    Utilities.sleep(4.02 * 1000);
}

function markChanges() {
  
  var ranges = buildRangesObject();
  var count = 0;
  
  if (Object.keys(ranges).length === 0) {
    throw new Error('No valid documentation sheets found. Remember to run the <strong>Build Documentation</strong> menu option first!');
  }

  for (var i in ranges) {
    
    var folders = ranges[i].folders.getValues().map(v => v[0]);
    var notes = ranges[i].notes.getValues().map(v => v[0]);
    var jsons = ranges[i].json.getValues().map(v => v[0]);
    var lookup = ranges[i].lookup_folders ? ranges[i].lookup_folders.getValues() : [];

    folders.forEach((folder, index) => {
      
      if (!folder.length)
        return;
      
      var cell = ranges[i].folders.getCell(index + 1, 1);
      var pid = JSON.parse(jsons[index]).parentFolderId || '';
      var fid = _lookup_folderid(folder, lookup);
     
      Logger.log('folder compare:', fid, pid);
      
      if (fid == pid) {
        cell.setBackground('#fff');
      } 
      else 
      if (fid != pid) {
        cell.setBackground('#fce5cd');
        count++;
      }
    });
    
    notes.forEach((note, index) => {
      
      var cell = ranges[i].notes.getCell(index + 1, 1);
      var json = JSON.parse(jsons[index]).notes || '';

      if (note == json) {
        cell.setBackground('#fff');
      } 
      else 
      if (note != json) {
        cell.setBackground('#fce5cd');
        count++;
      }
    });
    
  } 
    
  return count;
}

function processChanges(action) {
  
  var ranges = buildRangesObject();
  var toUpdate = [];
  var selectedAccountId = getAccountIdFromApiPath();
  var selectedContainerId = getContainerIdFromApiPath();

  for (var i in ranges) {
    
    var range = ranges[i];
    
    if (selectedAccountId != range.accountId || selectedContainerId != range.containerId)
      continue;
    
    var ids = range.ids.getValues().map(v => v[0]);
    var folders = range.folders.getValues().map(v => v[0]);
    var notes = range.notes.getValues().map(v => v[0]);
    var jsons = range.json.getValues().map(v => v[0]);
    var lookup = ranges[i].lookup_folders ? ranges[i].lookup_folders.getValues() : [];
    
    folders.forEach((folder, index) => {
                    
      /* won't allow removing folders for now */
      if (!folder.length)
        return;
    
      var cell = ranges[i].folders.getCell(index + 1, 1);
      var fid = _lookup_folderid(folder, lookup);
      var json = JSON.parse(jsons[index]);
      var pid = json.parentFolderId || '';
    
      if (fid == pid) {
        cell.setBackground('#fff');
        return;
      } 
    
      cell.setBackground('#fff');
    
      toUpdate.push({
        id: ids[index],
        name: json.name || '',
        type: range.type,
        accountId: range.accountId,
        containerId: range.containerId,
        fid: fid
      });
    
    });
  
    notes.forEach((note, index) => {
      
      var cell = range.notes.getCell(index + 1, 1);
      var json = JSON.parse(jsons[index]);
      
      json.notes = json.notes || '';
     
      if (note == json.notes) {
        cell.setBackground('#fff');
        return;
      }

      cell.setBackground('#fff');

      toUpdate.push({
          id: ids[index],
          name: json.name || '',
          type: range.type,
          accountId: range.accountId,
          containerId: range.containerId,
          notes: note
      });
    
    });

  }
  
  return toUpdate;
}

_lookup_folderid = (name, lookup) => {
  var match = lookup.filter(f => f[1] == name ? true : false)[0];
  return match ? parseInt(match[0]) : null;
};
  
_lookup_triggers = (tids, o) => {
  
  return tids.map(tid => {
               
    var t = o.triggers.filter(t => t.triggerId == tid)[0];
  
    if (!t || !t.triggerId)
      return tid;
  
    var tr = JSON.parse(JSON.stringify(t));
    delete tr.triggerId, delete tr.accountId, delete tr.containerId, delete tr.fingerprint, delete tr.parentFolderId, delete tr.workspaceId, delete tr.tagManagerUrl, delete tr.path;
  
    //return JSON.stringify(tr);
    return tr.name + ' (' + tr.type + ')';
  });
};

_lookup_folder = (flid, o) => {
  
  var fl = o.folders.filter(f => f.folderId == flid)[0];
  
  return fl ? fl.name : flid;
}

function formatTags(o) {
      
  var data = [];

  o.tags.forEach(function(tag) {

    var ta = JSON.parse(JSON.stringify(tag));
    delete ta.tagId, delete ta.type, delete ta.accountId, delete ta.containerId, delete ta.fingerprint, delete ta.path, delete ta.firingTriggerId, delete ta.blockingTriggerId, delete ta.setupTag, delete ta.teardownTag, delete ta.tagManagerUrl;
    
    data.push([
      tag.name,
      tag.tagId,
      tag.type,
      tag.parentFolderId ? _lookup_folder(tag.parentFolderId, o) : '',
      new Date(parseInt(tag.fingerprint)),
      tag.firingTriggerId ? _lookup_triggers(tag.firingTriggerId, o).join(', ') : '',
      tag.blockingTriggerId ? _lookup_triggers(tag.blockingTriggerId, o).join(', ') : '',
      tag.setupTag ? tag.setupTag[0].tagName : '',
      tag.teardownTag ? tag.teardownTag[0].tagName : '',
      tag.notes || '',
      JSON.stringify(ta)
    ]);
    
  });
  return data;
}

function formatVariables(o) {
  
  var data = [];
  
  o.variables.forEach(function(variable) {
    
    var v = JSON.parse(JSON.stringify(variable));
    delete v.variableId, delete v.accountId, delete v.containerId, delete v.fingerprint;
    
    data.push([
      variable.name,
      variable.variableId,
      variable.type,
      variable.parentFolderId ? _lookup_folder(variable.parentFolderId, o) : '',
      new Date(parseInt(variable.fingerprint)),
      variable.notes || '',
      JSON.stringify(v)
    ]);
    
  });
  return data;
}

function formatTriggers(o) {
  
  var data = [];
  
  o.triggers.forEach(function(trigger) {
    
    var tr = JSON.parse(JSON.stringify(trigger));
    delete tr.triggerId, delete tr.accountId, delete tr.containerId, delete tr.fingerprint, delete tr.workspaceId, delete tr.tagManagerUrl, delete tr.path;
    
    data.push([
      trigger.name,
      trigger.triggerId,
      trigger.type,
      trigger.parentFolderId ? _lookup_folder(trigger.parentFolderId, o) : '',
      new Date(parseInt(trigger.fingerprint)),
      trigger.notes || '',
      JSON.stringify(tr)
    ]);
  });
  return data;
}

function clearInvalidRanges() {
  
  var storedRanges = JSON.parse(PropertiesService.getUserProperties().getProperty('named_ranges')) || {};
  var storedRangesNames = Object.keys(storedRanges);
  
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  var namedRangesNames = namedRanges.map(function(a) { return a.getName(); });
  
  storedRangesNames.forEach(function(storedRangeName) {
    if (namedRangesNames.indexOf(storedRangeName) === -1) {
      SpreadsheetApp.getActiveSpreadsheet().removeNamedRange(storedRangeName);
      delete storedRanges[storedRangeName];
    }
  });
  
  PropertiesService.getUserProperties().setProperty('named_ranges', JSON.stringify(storedRanges));
}

function setNamedRanges(sheet, prefix, items, colLength) {
  
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var user_ranges = JSON.parse(PropertiesService.getUserProperties().getProperty('named_ranges')) || {};
  
  items.forEach(item => {
  
       var range = sheet.getRange(2, item.index, colLength, 1);
       var name = prefix + item.suffix;
       active.setNamedRange(name, active.getRange(sheet.getName() + '!' + range.getA1Notation()));
       user_ranges[name] = true;
  });

  PropertiesService.getUserProperties().setProperty('named_ranges', JSON.stringify(user_ranges));
}

function createHeaders(sheet, labels) {

  var labelsRange = sheet.getRange(1,1,1,labels.length);
  labelsRange.setValues([labels]).setFontWeight('bold').setFontColor('white').setBackground('#1155cc').setHorizontalAlignment('center');
}

function buildTriggerSheet(o, mode) {
  
  var sheetName = o.containerPublicId + '_triggers';
  var sheet = insertSheet(sheetName, mode);
  
  if (sheet === false) { return; }

  var triggerLabels = ['Trigger Name', 'Trigger ID', 'Trigger Type', 'Folder ID', 'Last Modified', 'Notes', 'JSON'];

  createHeaders(sheet, triggerLabels);

  var triggersObject = formatTriggers(o);
  
  if (!triggersObject.length)
    return;
    
  var dataRange = sheet.getRange(2, 1, triggersObject.length,triggerLabels.length);
  
  dataRange.setValues(triggersObject).setFontColor('black').setBackground('#fff').setFontWeight('normal');
  
  var prefix = 'triggers_' + o.accountId + '_' + o.containerId;
  
  setNamedRanges(sheet, prefix,  [{ index: 2, suffix: '_ids' }, { index: 4, suffix: '_folders' }, { index: 6, suffix: '_notes' }, { index: 7, suffix: '_json' }], triggersObject.length);
  
  var formats = triggersObject.map(function(a) {
    return ['@', '#', '@', '@', 'dd/mm/yy at h:mm', '@', '@'];
  });
  
  dataRange.setNumberFormats(formats);
  dataRange.setHorizontalAlignment('left');
  sheet.autoResizeColumns(1, 7);
  sheet.sort(5, false);
  
  if (o.folders) {
    var folder = SpreadsheetApp.newDataValidation().requireValueInList(o.folders.map(f => f.name)).build();
    sheet.getRange('D2:D').setDataValidation(folder);
    sheet.getRange('D1').setDataValidation(null);
  }
}

function buildVariableSheet(o, mode) {
  
  var sheetName = o.containerPublicId + '_variables';
  var sheet = insertSheet(sheetName, mode);
  
  if (sheet === false) { return; }

  var varLabels = ['Variable Name', 'Variable ID', 'Variable Type', 'Folder ID', 'Last Modified', 'Notes', 'JSON'];

  createHeaders(sheet, varLabels);

  var varObject = formatVariables(o);
  
  if (!varObject.length)
    return;
    
  var dataRange = sheet.getRange(2, 1, varObject.length, varLabels.length);
  dataRange.setValues(varObject).setFontColor('black').setBackground('#fff').setFontWeight('normal');
  
  var prefix = 'variables_' + o.accountId + '_' + o.containerId;
  
  setNamedRanges(sheet, prefix, [{ index: 2, suffix: '_ids' }, { index: 4, suffix: '_folders' }, { index: 6, suffix: '_notes' }, { index: 7, suffix: '_json' }], varObject.length);
  
  var formats = varObject.map(function(a) {
    return ['@', '#', '@', '@', 'mm/dd/yy at h:mm', '@', '@'];
  });
    
  dataRange.setNumberFormats(formats).setHorizontalAlignment('left');
  sheet.autoResizeColumns(1, 6);
  sheet.sort(5, false);

  
  if (o.folders) {
    var folder = SpreadsheetApp.newDataValidation().requireValueInList(o.folders.map(f => f.name)).build();
    sheet.getRange('D2:D').setDataValidation(folder);
    sheet.getRange('D1').setDataValidation(null);
  }
}

function buildTagSheet(o, mode) {
  
  var sheetName = o.containerPublicId + '_tags';
  var sheet = insertSheet(sheetName, mode);
  
  if (sheet === false) { return; }

  var tagLabels = ['Tag Name', 'Tag ID', 'Tag Type', 'Folder ID', 'Last Modified', 'Firing Triggers', 'Blocking Triggers', 'Setup Tag', 'Cleanup Tag', 'Notes', 'JSON'];

  createHeaders(sheet, tagLabels);

  //SpreadsheetApp.getUi().alert('Tags: ' + JSON.stringify(o.tags));
  
  var clip = SpreadsheetApp.WrapStrategy.CLIP;
  sheet.getRange("F:G").setWrapStrategy(clip);
  sheet.getRange("K:K").setWrapStrategy(clip);
  
  var tagsObject = formatTags(o);
  
  if (!tagsObject.length)
    return;
    
  var dataRange = sheet.getRange(2, 1, tagsObject.length, tagLabels.length);
  dataRange.setValues(tagsObject).setFontColor('black').setBackground('#fff').setFontWeight('normal');
  
  var prefix = 'tags_' + o.accountId + '_' + o.containerId;
  setNamedRanges(sheet, prefix, [{ index: 2, suffix: '_ids' }, { index: 4, suffix: '_folders' }, { index: 10, suffix: '_notes' }, { index: 11, suffix: '_json' }], tagsObject.length);
  
  var formats = tagsObject.map(function(a) {
    return ['@', '#', '@', '@', 'mm/dd/yy at h:mm', '@', '@', '@', '@', '@', '@'];
  });
  
  dataRange.setNumberFormats(formats).setHorizontalAlignment('left');
  
  sheet.autoResizeColumns(1, 11);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 300);
  sheet.setColumnWidth(11, 300);
  sheet.sort(5, false);
  
  if (o.folders) {
    var folder = SpreadsheetApp.newDataValidation().requireValueInList(o.folders.map(f => f.name)).build();
    sheet.getRange('D2:D').setDataValidation(folder);
    sheet.getRange('D1').setDataValidation(null);
  }
}

function buildContainerSheet(o, index, latest, mode) {
  
  var sheetName = o.containerPublicId + '_container';
  var sheet = index ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) : insertSheet(sheetName, mode);
  
  if (sheet === false) { return; }
  
  if (index) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 4);
      sheet.getRange("A7:B").moveTo(sheet.getRange("A11"));
  }

  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 340);
  
  var containerHeader = sheet.getRange(1,1,1,2);
  containerHeader.setValues([['Google Tag Manager Documentation','']]);
  containerHeader.mergeAcross();
  containerHeader.setBackground('#1155cc');
  containerHeader.setFontWeight('bold');
  containerHeader.setHorizontalAlignment('center');
  containerHeader.setFontColor('white');
  
  var containerLabels = ['Container ID:', 'Container Name:', 'Container Notes:', 'Version ID:', 'Version Name:', 'Version Description:', 'Published:', 'Container Link:', 'API Path:'];
  
  var containerContent = sheet.getRange(2, 1, containerLabels.length, 2);
  
  containerContent.setValues([
    [containerLabels[0], o.containerPublicId],
    [containerLabels[1], o.containerName],
    [containerLabels[2], o.containerNotes],
    [containerLabels[7], o.containerLink],
    [containerLabels[8], 'accounts/' + o.accountId + '/containers/' + o.containerId + '/versions/' + o.versionId],
    [containerLabels[3], o.versionId],
    [containerLabels[4], o.versionName],
    [containerLabels[5], o.versionDescription],
    [containerLabels[6], o.versionCreatedOrPublished],
  ]);
 
  if (index == 0)
    sheet.getRange("B5:B6").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    
  containerContent.setBackground('white').setFontColor('black');
  sheet.getRange("A7:B7").setBorder(true, null, null, null, null, null);
  sheet.getRange("B9").setWrap(true);
  sheet.getRange("B10").setNumberFormat('mm/dd/yy at h:mm');
  containerContent.setVerticalAlignment('top');
  sheet.getRange("B7:B10").setFontWeight(latest ? 'bold' : 'normal');
    
  var containerLabelCol = sheet.getRange(2,1,containerLabels.length,1);
  containerLabelCol.setFontWeight('bold');
  containerLabelCol.setHorizontalAlignment('right');
  
  var containerDataCol = sheet.getRange(2,2,containerLabels.length,1);
  containerDataCol.setHorizontalAlignment('left');

  if (!latest) 
    return;
  
  var overviewHeader = sheet.getRange(1,4,1,8);
  overviewHeader.setValues([['Contents Overview v' + o.versionId, '', '', '', '', '', '', '']]);
  overviewHeader.mergeAcross();
  overviewHeader.setBackground('#85200c');
  overviewHeader.setFontWeight('bold');
  overviewHeader.setHorizontalAlignment('center');
  overviewHeader.setFontColor('white');
  
  var overviewSubHeader = sheet.getRange(2,4,1,8);
  overviewSubHeader.setValues([['Tag Type', 'Quantity', 'Trigger Type', 'Quantity', 'Variable Type', 'Quantity', 'Folder ID', 'Folder name']]);
  overviewSubHeader.setHorizontalAlignments([['right','left','right','left','right','left', 'right', 'left']]);
  overviewSubHeader.setFontWeight('bold');
  overviewSubHeader.setBackground('#e6d6d6');
  
  var tags = getAssetOverview(o.tags);
  var tagsRange = sheet.getRange(3,4,tags.sortedlist.length,2);
  var tagsSum = tags.sum;
  
  tagsRange.setValues(tags.sortedlist);
  sheet.getRange(3,4,tags.sortedlist.length,1).setHorizontalAlignment('right');
  sheet.getRange(3,5,tags.sortedlist.length,1).setHorizontalAlignment('left');

  var triggers = getAssetOverview(o.triggers);
  var triggersRange = sheet.getRange(3,6,triggers.sortedlist.length,2);
  var triggersSum = triggers.sum;
  
  triggersRange.setValues(triggers.sortedlist);
  sheet.getRange(3,6,triggers.sortedlist.length,1).setHorizontalAlignment('right');
  sheet.getRange(3,7,triggers.sortedlist.length,1).setHorizontalAlignment('left');

  var variables = getAssetOverview(o.variables);
  var variablesRange = sheet.getRange(3,8,variables.sortedlist.length,2);
  var variablesSum = variables.sum;
  
  variablesRange.setValues(variables.sortedlist);
  sheet.getRange(3,8,variables.sortedlist.length,1).setHorizontalAlignment('right');
  sheet.getRange(3,9,variables.sortedlist.length,1).setHorizontalAlignment('left');
  
  var folders = o.folders.map(function(folder) {
    return [folder.folderId, folder.name];
  });
  
  if (folders.length) {
    var foldersRange = sheet.getRange(3, 10, folders.length, 2);
    foldersRange.setValues(folders);
    SpreadsheetApp.getActiveSpreadsheet().setNamedRange('folders_' + o.accountId + '_' + o.containerId, sheet.getRange(sheet.getName() + '!' + foldersRange.getA1Notation()));
  }
  
  var contentLength = Math.max(tags.sortedlist.length, variables.sortedlist.length, triggers.sortedlist.length, folders.length);
  var totalRow = sheet.getRange(contentLength + 3, 4, 1, 8);
  totalRow.setValues([
    ['Total Tags:', tagsSum, 'Total Triggers:', triggersSum, 'Total Variables:', variablesSum, '', '']
  ]);
  totalRow.setHorizontalAlignments([['right', 'left', 'right', 'left', 'right', 'left', 'right', 'left']]);
  totalRow.setFontWeight('bold');
  //totalRow.setBackground('#e6d6d6');
  
  sheet.autoResizeColumns(4, 8);
}

function startProcess(aid, cid, version_count, mode) {

  mode = mode || 'build';
  
  Logger.log('startProcess:', aid, cid, version_count, mode);
  
  var versions = fetchVersions(aid, cid).reverse();
  
  version_count = version_count == 'all' ? versions.length : parseInt(version_count);

  versions = versions.slice(0, version_count).reverse();
  
  for (var i = 0; i < versions.length; i++) {

    var v = fetchVersion(aid, cid, versions[i].containerVersionId);
    var latest = (i == (versions.length - 1));
    
    var o = {
      accountId: v.container.accountId,
      containerId: v.container.containerId,
      containerName: v.container.name,
      containerPublicId: v.container.publicId,
      containerNotes: v.container.notes || '',
      containerLink: v.container.tagManagerUrl,
      versionName: v.name || '',
      versionId: v.containerVersionId,
      versionDescription: v.description || '',
      versionCreatedOrPublished: new Date(parseInt(v.fingerprint)),
      tags: v.tag || [],
      variables: v.variable || [],
      triggers: v.trigger || [],
      folders: v.folder || []
    };
    
    buildContainerSheet(o, i, latest, mode);
    
    if (latest) {
      
      o.triggers = fetchTriggers(aid, cid);
     
      buildTagSheet(o, mode);
      buildTriggerSheet(o, mode);
      buildVariableSheet(o, mode);
    }
  }
  
}

quickRebuild = () => {
  
  var path = getApiPath();
  var version_count = getVersionCount() || 5;
  
  if (!path)
    return SpreadsheetApp.getUi().alert('For "Quick Rebuild" to work, you need to be in a tab created by "Build Documentation" and the respective _container tab with an API Path field needs to be populated.');
  
  startProcess(getAccountIdFromApiPath(path), getContainerIdFromApiPath(path), version_count, 'rebuild');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function openContainerSelector() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile('ContainerSelector').evaluate().setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Build Documentation');
}

function openMarkChangesModal() {
  clearInvalidRanges();
  var ui = SpreadsheetApp.getUi();
  if (Object.keys(buildRangesObject()).length === 0) {
    ui.alert('No valid documentation sheets found! Run "Build Documentation" if necessary.');
    return;
  }
  var html = HtmlService.createTemplateFromFile('MarkChangesModal').evaluate().setWidth(500).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mark Changes to Folders & Notes');
}

function openPushChangesModal() {
  clearInvalidRanges();
  var ui = SpreadsheetApp.getUi();
  if (getApiPath() === false) {
    ui.alert('You need to have a valid documentation sheet selected first! Run "Build Documentation" if necessary.', ui.ButtonSet.OK);
    return;
  }
  var html = HtmlService.createTemplateFromFile('PushChangesModal').evaluate().setWidth(500).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Push Changes to Folders & Notes');
}

function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Build Documentation', 'openContainerSelector')
  menu.addItem('Quick Rebuild', 'quickRebuild')
  menu.addItem('Mark Changes', 'markChanges')
  menu.addItem('Push Changes', 'openPushChangesModal')
  menu.addToUi();
}