var win32ole = require("./lib/win32ole");

exports.action = function(data, callback, config, SARAH) {
	var activeXObject = win32ole.client.Dispatch("X10.ActiveHome")
	activeXObject.SendAction(data.type, data.code + " " + data.valeur);
	var tts = ''; 
  	callback({ 'tts': tts });
}