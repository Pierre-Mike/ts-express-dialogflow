var { cache } = require("../lib");
var graph = require("@microsoft/microsoft-graph-client");
const { Card } = require("dialogflow-fulfillment");

function getClientGraph(
	accessToken: String,
	apiUrl: String,
	search?: String,
	filter?: String,
	select?: String
): Promise<any> {
	return new Promise((resolve, reject) => {
		var client = graph.Client.init({
			authProvider: (callback: Function) => {
				callback(null, accessToken);
			}
		}).api(apiUrl);
		if (search) {
			client.search(search);
		}
		if (filter) {
			client.filter(filter);
		}
		if (select) {
			client.select(select);
		}
		client
			.get()
			.catch((err: any) => reject(err))
			.then((e: any) => resolve(e));
	});
}

function sharepointContext(agent: any) {
	console.log(agent.context);

	if (cache.get(agent.context.session, false)) {
		console.log("set from cache to context");
		agent.context.set({
			name: "sharepoint_connection",
			lifespan: 50,
			parameters: { access_token: cache.get(agent.context.session) }
		});
	}
	return agent.context.get("sharepoint_connection", {}).parameters;
}

export async function search(agent: any) {
	var cntxtParam = sharepointContext(agent);
	console.log("cntxtParam", cntxtParam);
	if ("access_token" in cntxtParam) {
		let res: any = await getClientGraph(
			cntxtParam.access_token,
			"/sites",
			agent.parameters.name
		);
		console.log(res);
		if (res.value.length === 0) {
			agent.add(`No site found for "${agent.parameters.name}".`);
		} else {
			res.value.map((e: any) =>
				agent.add(
					new Card({
						title: e.displayName,
						buttonText: e.name,
						buttonUrl: e.webUrl
					})
				)
			);
		}
	} else {
		agent.add("you are not connected !");
	}
}
