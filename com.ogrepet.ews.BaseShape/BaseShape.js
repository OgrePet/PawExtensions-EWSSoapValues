var BaseShape = function() {

  this.evaluate = function() {
    var shapeType = this.propsSet
    var dynamicValue = "<t:BaseShape>"+shapeType+"</t:BaseShape>"; // generate some dynamic value
    return dynamicValue;
  }

  this.text = function(context) {
  		if (!this.propsSet && this.propsSet.length === 0) {
			return null;
		}
		  var shapeType = this.propsSet
  		return ": " + shapeType
	}
}

BaseShape.identifier = "com.ogrepet.ews.BaseShape";
BaseShape.title = "EWS Base Shape";
BaseShape.help = "https://msdn.microsoft.com/en-us/library/office/aa580545(v=exchg.150).aspx";

BaseShape.inputs = [
    InputField("propsSet", "Properties Set", "Select",
     {"choices": {"IdOnly": "IdOnly", "Default": "Default", "AllProperties": "AllProperties"}, defaultValue : "IdOnly"})
];

registerDynamicValueClass(BaseShape)
