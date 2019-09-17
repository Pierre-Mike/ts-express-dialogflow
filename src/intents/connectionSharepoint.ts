const { Card } = require("dialogflow-fulfillment");
import { CONFIG, oauth2 } from "../lib";

function getUrlLogin(agent: any, auth: any, redirect_uri: String): String {
	let state = agent.context.session;
	console.log(state);
	var urlAuthorization = auth.authorizationCode.authorizeURL({
		redirect_uri,
		state
	});
	return urlAuthorization + "&resource=https://graph.microsoft.com/";
}

const setContextAgent = (agent: any, name: String, lifespan: Number): any => {
	agent.context.set({
		name,
		lifespan
	});
	return agent;
};

export const connection = (agent: any): void => {
	let card = new Card("Connection to SharePoint");
	let urlConnection = getUrlLogin(agent, oauth2, CONFIG.microsoft.redirect_uri);
	card.setImage(
		"http://www.tascmanagement.com/wp-content/uploads/2016/05/course-logo-small-SP-300x250.png"
	);
	card.setText(
		"This is the body text of a card.  You can even use line\nbreaks and emoji! 💁"
	);
	card.setButton({
		text: "Login",
		url: urlConnection
	});

	agent = setContextAgent(agent, "sharepoint_connection", 50);
	agent.add(card); // return agent
};
