var myDocument = app.activeDocument;
app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;

var alaboshuzi = new Array("0","1","2","3","4","5","6","7","8","9")
var zangshuzi = new Array("༠","༡","༢","༣","༤","༥","༦","༧","༨","༩")
for (n=0; n<10; n++){
    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = false;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;
    
    app.findTextPreferences.findWhat = zangshuzi[n];
    app.changeTextPreferences.changeTo = alaboshuzi[n];
    myDocument.selection[0].changeText();
    
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
}