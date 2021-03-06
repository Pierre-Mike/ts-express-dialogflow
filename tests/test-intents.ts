"use strict";

const chai = require("chai");
const spies = require("chai-spies");
const expect = chai.expect;

import { getMapIntents } from "../src/getIntents";
import { welcome } from "../src/intents/welcome";
import { reset } from "../src/intents/reset";
import { fallback } from "../src/intents/fallback";
import { connection } from "../src/intents/connectionSharepoint";

chai.use(spies);

describe("getMapIntents", () => {
	it("should have function for every intent", async () => {
		var intentsName = [
			"Default Welcome Intent",
			"Login SharePoint",
			"Search SharePoint",
			"Fallback",
			"reset"
		];
		var intentsMap = getMapIntents();
		// EVERY INTENT HAS TO BE IN THE INTENTMAP
		intentsName.forEach(e => expect(intentsMap.has(e)).to.be.true);
	});
});

describe("intents", () => {
	var spyAdd = chai.spy();
	it("welcome", async () => {
		welcome({ add: spyAdd });
		expect(spyAdd).called.with("hello express");
	});
	it("reset", async () => {
		var spyAdd = chai.spy();
		var spyDel = chai.spy();
		reset({ add: spyAdd, context: { delete: spyDel } });
		expect(spyAdd).called.with("reset successful");
		expect(spyDel).called.with("sharepoint_connection");
	});
	it("fallback", async () => {
		var spyAdd = chai.spy();
		fallback({ add: spyAdd });
		expect(spyAdd).called.with("NO INTENT");
	});
	it("connection", async () => {
		var spyAdd = chai.spy();
		var spySet = chai.spy();
		connection({ add: spyAdd, context: { set: spySet } });
		expect(spyAdd).called.once;
		expect(spySet).called.with({
			name: "sharepoint_connection",
			lifespan: 50
		});
	});
});
