# Code snippets are only available for the latest version. Current version is 1.x
from msgraph import GraphServiceClient
from msgraph.generated.models.external_connectors.external_connection import ExternalConnection
# To initialize your graph_client, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=python

from config import CONFIG

scopes = ['User.Read']

tenant_id = CONFIG['tenant_id']
client_id = CONFIG['client_id']

# azure.identity
credential = DeviceCodeCredential(
    tenant_id=tenant_id,
    client_id=client_id)

graph_client = GraphServiceClient(credential, scopes)


request_body = ExternalConnection(
	id = "project-portal-connection",
	name = "Project Portal Connector",
	description = "Connection to index Project Portal system",
)

result = await graph_client.external.connections.post(request_body)