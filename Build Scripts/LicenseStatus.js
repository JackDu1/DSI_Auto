importClass(java.lang.System);
importClass(java.util.regex.Matcher);
importClass(java.util.regex.Pattern);
importClass(java.util.logging.Level);
importClass(com.gargoylesoftware.htmlunit.FailingHttpStatusCodeException);
importClass(com.gargoylesoftware.htmlunit.NicelyResynchronizingAjaxController);
importClass(com.gargoylesoftware.htmlunit.WebClient);
importClass(com.gargoylesoftware.htmlunit.html.DomElement);
importClass(com.gargoylesoftware.htmlunit.html.HtmlPage);

java.util.logging.Logger.getLogger("com.gargoylesoftware").setLevel(Level.OFF);
System.out.println(">>>>>>>>>>>>>License Status<<<<<<<<<<<<<<<<<");
var strTARGET_URL=project.getProperty("licenseStatusURL");
var webClient = new WebClient();
webClient.getOptions().setJavaScriptEnabled(true);
webClient.getOptions().setCssEnabled(false);
webClient.setAjaxController(new NicelyResynchronizingAjaxController());
webClient.setJavaScriptTimeout(35000);
webClient.getOptions().setThrowExceptionOnScriptError(false);
var htmlPage = webClient.getPage(strTARGET_URL);
var domElement = htmlPage.getElementById("mydata");
var regEx = "TestExecute[\\w\\s-]*?Sessions|Product[\\w\\s-]*?Actions";
var pat = Pattern.compile(regEx);
var matcher = pat.matcher(domElement.asText());
while(matcher.find()){
	System.out.println(matcher.group().replaceAll("\r\n", ""));
}
webClient.closeAllWindows();

System.out.println("\n>>>>>>>>>>>>>License Session<<<<<<<<<<<<<<<<<");
strTARGET_URL=project.getProperty("licenseSessionURL");
htmlPage = webClient.getPage(strTARGET_URL);
domElement = htmlPage.getElementById("mydata");
regEx = "TestExecute[\\w\\s-:\\.,]*?Disconnect|Product[\\w\\s-]*?Actions";
pat = Pattern.compile(regEx);
matcher = pat.matcher(domElement.asText());
while(matcher.find()){
	System.out.println(matcher.group().replaceAll("\r\n", ""));
}
webClient.closeAllWindows();

