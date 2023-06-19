# Node-Red-Azure-nodes

This package provides a collection of custom nodes for interacting with Azure services in Node-RED.

## Nodes

### azure_get_access_token

This node retrieves an access token for authenticating Azure API requests. 
It requires the `msg.config.Azure.clientId`, `msg.config.Azure.clientSecret`, and `msg.config.Azure.tenantId` properties to be set. 
The access token is then available in `msg.access_token` and can be used by other nodes.

### azure_users_to_id

This node retrieves the Azure ID of a user based on their email address. 
It requires the `msg.access_token` and `msg.emails` properties to be set. 
The node returns the user's ID in `msg.usersId`.

### azure_does_group_exist

This node checks if a group exists in Azure Active Directory based on its name. 
It requires the `msg.group_name` and `msg.access_token` properties to be set. 
If the group exists, its ID is returned in `msg.groupID`. If there are multiple groups with the same name, an exception is thrown.

### azure_create_group

This node creates a security group in Azure Active Directory. 
It requires the `msg.group_name`, `msg.group_desc`, and `msg.access_token` properties to be set. 
The created group will have an empty `groupTypes`, `mailEnabled` set to `false`, and `securityEnabled` set to `true`. 
The group's name and description will be provided by the client, and its `mailNickname` will be a concatenation of the string 'Authomize_' and the `msg.group_name`. 
The node returns the group's ID in `msg.groupID`.

### azure_is_user_in_group

This node checks if a user is a member of a group in Azure Active Directory. 
It requires the `msg.access_token`, `msg.usersId`, `msg.emails`, `msg.group_name`, and `msg.groupID` properties to be set. 
This node acts as a switch, passing the `msg` object through different outputs depending on the response.

### azure_add_user_to_group

This node adds users to a group in Azure Active Directory. 
It requires the `msg.access_token`, `msg.usersId`, `msg.emails`, `msg.group_name`, and `msg.groupID` properties to be set. 
The node returns the original `msg` object without modifying it and prints a message indicating that the user has been added to the group.

### azure_remove_user_from_group

This node removes users from a group in Azure Active Directory. 
It requires the `msg.access_token`, `msg.usersId`, `msg.emails`, `msg.group_name`, and `msg.groupID` properties to be set. 
The node returns the original `msg` object without modifying it and prints a message indicating that the user has been removed from the group.

### azure_disable_user

This node disables a user in Azure Active Directory. 
It requires the `msg.access_token`, `msg.emails`, and `msg.usersId` properties to be set. 
The node returns the original `msg` object without modifying it and prints a message indicating that the user has been disabled in Azure.

## Usage

1. Install the node-red-azure-nodes package using npm: npm i node-red-azure-nodes

2. In your Node-RED flow, you will find the Azure nodes under the "Authomize Azure" category in the Node-RED palette.

3. Drag and drop the desired nodes into your flow and configure their properties as required.

4. Connect the nodes to create the desired workflow.

## Contributing

Contributions, bug reports, and feature requests are welcome!

## License

This project is licensed under the [MIT License](LICENSE).