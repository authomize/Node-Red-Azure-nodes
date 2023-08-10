const axios = require('axios');

module.exports = function(RED) {
    function azure_users_to_id_Node(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        var results = [];
        node.auth = RED.nodes.getNode(config.auth);
        
        node.on('input', async function(msg) {
            if (!node.auth || !node.auth.has_credentials) {
				node.error("auth configuration is missing");
				return
			}

			const access_token = await node.auth.get_access_token();
			const emails = msg.emails || (msg.payload.data.entities[0] ? [msg.payload.data.entities[0].email] : []);
			
            try {
               	const usersId = [];

				for (const email of emails) {
					var searchURL = 'https://graph.microsoft.com/v1.0/users?$filter=mail%20eq%20%27' + email + '%27%20or%20userPrincipalName%20eq%20%27' + email + '%27';
					
					const response = await axios.get(searchURL, {
						headers: {
						  Authorization: 'Bearer ' + access_token
						}
				  });

					if (response.data.value.length == 1) {
						usersId.push(response.data.value[0].id);
						msg.usersId = usersId;			
					} else if (response.data.value.length > 1){
						throw new Error('There is more than one user with the suggested name');
					}else {
						throw new Error('Current user ' + email + ' Doesn\'t exist in Azure');
				  }
				}

                node.send(msg);				

            } catch (error) {
                node.warn(error);
            }
        });
    }

    RED.nodes.registerType("azure_users_to_id", azure_users_to_id_Node);
};
