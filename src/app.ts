import express = require("express");
const { WebhookClient } = require("dialogflow-fulfillment");
const bodyParser = require("body-parser");
const { CONFIG, cache, oauth2 } = require("./lib");

const PORT = process.env.PORT || 8080;
const intents = require("./getIntents");
const basicAuth = require("express-basic-auth");

// Create a new express application instance
const app: express.Application = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.get("/", function(req, res) {
	res.send("Hello World!");
});

app.get("/getAToken", async (req, res) => {
	const options = {
		code: req.query.code,
		redirect_uri: CONFIG.microsoft.redirect_uri
	};
	const result = await oauth2.authorizationCode.getToken(options);
	const access_token = oauth2.accessToken.create(result).token.access_token;

	cache.set(req.query.state, access_token, 100000000);
	console.log("store token into the cache");
	res.send(JSON.stringify(cache.get(req.query.state)));
});

app.post("/", basicAuth(CONFIG.auth), (req, res) => {
	console.info(`\n\n>>>>>>> S E R V E R   H I T <<<<<<<`);
	const agent = new WebhookClient({ req, res });
	agent.handleRequest(intents());
});

app.listen(3000, function() {
	console.log("Example app listening on port 3000 !");
});
