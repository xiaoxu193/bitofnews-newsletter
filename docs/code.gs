var ui = DocumentApp.getUi();
var data;

function onOpen() {
  ui.createMenu('Bit of News')
      .addItem('Render HTML', 'showSidebar')
      .addItem('Preview', 'showPreview')
      .addToUi();
}

function showPreview(newmember) {
  newmember = typeof newmember !== 'undefined' ? newmember : false;
  if (data == null)
    data = parseDoc();
  var newsletter = HtmlService.createTemplateFromFile('dailynews');
  newsletter.data = data;
  newsletter.newmember = newmember;
  newsletter.worldCupDays = getWorldCupDays();
  newsletter = newsletter
               .evaluate().setWidth(800).setHeight(600);
  ui.showModalDialog(newsletter, 'Bit of News');
}

function showSidebar(newmember) {
  newmember = typeof newmember !== 'undefined' ? newmember : false;
  var html = HtmlService.createTemplateFromFile('sidebar');
  data = parseDoc();
  html.data = data;
  html.newmember = newmember;
  html.worldCupDays = getWorldCupDays();
  
  html = html
         .evaluate()
         .setTitle('Render newsletter HTML').setWidth(300);

  ui.showSidebar(html);
}



//parse document into JSON object containing summaries, others, quote, and tidbit
function parseDoc(){
  var sections = getSections();
  
  var data = {"quote":"", "summaries":[], "others":[], "tidbit":"", "bottomquote":""};
  
  for (i in sections)
  {
    var sect = sections[i];
    var type = getType(sect);
        
    if (type == "_SUMMARY")
    {
      var summary = {};
      
      [summary.publisher, summary.category] = sect[0].getText().split(', ');
      summary.title = sect[1].getText();
      summary.link = sect[2].getText();
      summary.summary = sect[3].getText();
      //parse publisher and stuff
      data.summaries.push(summary);
    }
    else if (type == "_QUOTE")
    {
      var quote = {};
      quote.quote = sect[1].getText();
      
      description = sect[2].editAsText().copy();
      var format_indices = description.getTextAttributeIndices();
      Logger.log('quote indices');
      Logger.log(format_indices);
      if (format_indices.length ==2)
      {
        description.insertText(format_indices[1], '</a>');
        description.insertText(format_indices[0], '<a href="' + description.getLinkUrl(format_indices[0]) + '">');
      }
      quote.description = description.getText()
      data.quote = quote;
    }
    else if (type == "_OTHER")
    {
      var other = {};
      var text = sect[0].editAsText().copy();
      var format_indices = text.getTextAttributeIndices();
      if (format_indices.length !=4)
        continue;
      Logger.log(format_indices);
      text.insertText(format_indices[3], '</a>');
      text.insertText(format_indices[2], '<a href="' + text.getLinkUrl(format_indices[2]) + '">');
      var tail = text.getText().substring(format_indices[1]);
      
      var head = text.getText().substring(0,format_indices[1]);
      
      other.head = head;
      other.tail = tail;
      data.others.push(other);
    }
    else if (type == "_TIDBIT")
    {
      var tidbit = {};
      tidbit.title = sect[1].getText();
      
      
      tidsummary = sect[2].editAsText().copy();
      var format_indices = tidsummary.getTextAttributeIndices();
      Logger.log('tidbit');
      Logger.log(format_indices);
      if (format_indices.length ==3)
      {
        tidsummary.insertText(format_indices[2], '</a>');
        tidsummary.insertText(format_indices[1], '<a href="' + tidsummary.getLinkUrl(format_indices[1]) + '">');
      }
      tidbit.summary = tidsummary.getText()
      
      
      data.tidbit = tidbit;
    }
    else if (type == "_BOTTOMQUOTE")
    {
      var bottomquote = sect[1].getText();
      data.bottomquote = bottomquote;
    }
  }
  
  return data;
}

//returns the type of paragraph
function getType(sect){
  if (sect.length == 4)
      return "_SUMMARY";
  else if (sect.length == 3){
      var par = sect[0];
      if (par.findText('/QUOTE/'))
        return "_QUOTE";
      else if (par.findText('/TIDBIT/'))
        return "_TIDBIT";
  }
  else if (sect.length == 2) {
       var par = sect[0];
       if (par.findText('/BOTTOMQUOTE/'))
         return "_BOTTOMQUOTE";
  }
  else if (sect.length == 1)
       return "_OTHER";
}

//separate paragraphs into sections, delimited by empty paragraphs
function getSections(){
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paras = body.getParagraphs();
  
  var sections = [];
  var section = [];
  
  for (i in paras){
    var para = paras[i];
    
    if (!isEmpty(para) && i<paras.length-1){
      section.push(para);
    }
    else{
      if (section.length>0)
        sections.push(section);
      section = [];
    }
  }
  return sections;
}

//check if a paragraph is empty
function isEmpty(paragraph){
  if (paragraph.getText() =="")
    return true;
  else
    return false;
}

function getWorldCupDays(){
  var today = new Date();
  var start = new Date(2014, 5, 11, 12, 0, 0, 0);
  return Math.round((today-start)/(1000*60*60*24));
  
}