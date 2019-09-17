import { fallback } from "./intents/fallback";
import { search } from "./intents/searchSharepoint";
const { connection } = require("./intents/connectionSharepoint");
const welcome = require("./intents/welcome");
const reset = require("./intents/reset");

export function getMapIntents() {
	let intentMap = new Map();
	intentMap.set("Default Welcome Intent", welcome);
	intentMap.set("Login SharePoint", connection);
	intentMap.set("Search SharePoint", search);
	intentMap.set("Fallback", fallback);
	intentMap.set("reset", reset);
	return intentMap;
}
