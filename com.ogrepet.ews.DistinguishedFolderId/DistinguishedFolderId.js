var DistinguishedFolderId = function() {

  this.evaluate = function() {
    var shapeType = this.propsSet
    var dynamicValue = "<t:DistinguishedFolderId Id=\""+shapeType+"\"></t:DistinguishedFolderId>";

    return dynamicValue;
  }

  this.text = function(context) {
  		if (!this.propsSet && this.propsSet.length === 0) {
			return null;
		}
		  var shapeType = this.propsSet
  		return ": " + shapeType
	}

// 	return this
}

DistinguishedFolderId.identifier = "com.ogrepet.ews.DistinguishedFolderId";
DistinguishedFolderId.title = "EWS Distinguished Folder Id";
DistinguishedFolderId.help = "https://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx";

DistinguishedFolderId.inputs = [
    InputField("propsSet", "Properties Set", "Select",
     {"choices": {"calendar": "calendar",
     "contacts": "contacts",
     "deleteditems": "deleteditems",
     "drafts": "drafts",
     "inbox": "inbox",
     "journal": "journal",
     "notes": "notes",
     "outbox": "outbox",
     "sentitems": "sentitems",
     "tasks": "tasks",
     "msgfolderroot": "msgfolderroot",
     "root": "root",
     "junkemail": "junkemail",
     "searchfolders": "searchfolders",
     "voicemail": "voicemail",
     "recoverableitemsroot": "recoverableitemsroot",
     "recoverableitemsdeletions": "recoverableitemsdeletions",
     "recoverableitemsversions": "recoverableitemsversions",
     "recoverableitemspurges": "recoverableitemspurges",
     "archiveroot": "archiveroot",
     "archivemsgfolderroot": "archivemsgfolderroot",
     "archivedeleteditems": "archivedeleteditems",
     "archiveinbox": "archiveinbox",
     "archiverecoverableitemsroot": "archiverecoverableitemsroot",
     "archiverecoverableitemsdeletions": "archiverecoverableitemsdeletions",
     "archiverecoverableitemsversions": "archiverecoverableitemsversions",
     "archiverecoverableitemspurges": "archiverecoverableitemspurges",
     "syncissues": "syncissues",
     "conflicts": "conflicts",
     "localfailures": "localfailures",
     "serverfailures": "serverfailures",
     "recipientcache": "recipientcache",
     "quickcontacts": "quickcontacts",
     "conversationhistory": "conversationhistory",
     "adminauditlogs": "adminauditlogs",
     "todosearch": "todosearch",
     "mycontacts": "mycontacts",
     "directory": "directory",
     "imcontactlist": "imcontactlist",
     "peopleconnect": "peopleconnect",
     "favorites": "favorites"}, defaultValue : "notes"})
];


registerDynamicValueClass(DistinguishedFolderId)
