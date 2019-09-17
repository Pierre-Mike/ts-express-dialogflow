export function reset(agent: any): void {
	console.log(agent);
	agent.context.delete("sharepoint_connection");
	agent.add("reset successful");
}
