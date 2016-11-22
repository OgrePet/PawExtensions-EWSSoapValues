var ExtendedFieldURI_Tag = function() {

  this.evaluate = function() {
    if (!this.tag && this.tag.length === 0) {return "";}
    if (!this.type && this.type.length === 0) {return "";}

    var dynamicValue = "<t:ExtendedFieldURI PropertyTag=\""+this.tag+"\" PropertyType=\"" + this.type+"\"/>"; // generate some dynamic value
    return dynamicValue;
  }

  this.text = function(context) {
  		if (!this.tag && this.tag.length === 0) {return null;}
      if (!this.type && this.type.length === 0) {return null;}

		  var shapeType = this.propsSet
  		return ": " + this.tag + " (" + this.type + ")"
	}
}

ExtendedFieldURI_Tag.identifier = "com.ogrepet.ews.ExtendedFieldURI_Tag";
ExtendedFieldURI_Tag.title = "EWS ExtendedFieldURI (Tag)";
ExtendedFieldURI_Tag.help = "https://msdn.microsoft.com/en-us/library/office/aa564843(v=exchg.150).aspx";

ExtendedFieldURI_Tag.inputs = [
  InputField("tag", "Property Tag", "String"),
  InputField("type", "Property Type", "String", {defaultValue: "Binary"}),
];

registerDynamicValueClass(ExtendedFieldURI_Tag)
