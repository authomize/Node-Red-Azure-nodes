const axios = require('axios');

module.exports = function(RED) {
    function azure_add_user_to_group_Node(config) {
        RED.nodes.createNode(this, config);
        var node = this;

        node.auth = RED.nodes.getNode(config.auth);
       
        node.on('input', async function(msg) {

            if (!node.auth || !node.auth.has_credentials) {
				node.error("auth configuration is missing");
				return
			}

			const access_token = await node.auth.get_access_token();

			const groupId = RED.util.evaluateNodeProperty(
				config.groupId, config.groupIdType, node, msg
			  )
		
			  const userId = RED.util.evaluateNodeProperty(
				config.userId, config.userIdType, node, msg
			  )

			try {
				var url = `https://graph.microsoft.com/v1.0/groups/${groupId}/members/$ref`;
				
				const postData = {
					"@odata.id": 'https://graph.microsoft.com/v1.0/directoryObjects/' + userId
				};
				
				const headers = { Authorization: 'Bearer ' + access_token, "Content-Type": 'application/json' };
            
                axios.post(url, postData, { headers }).then(response => {
					node.warn(' user id: ' + userId + ' was added group Id ' + groupId);
					node.send(msg);
                }).catch(error => {
					node.warn('!!!!!!!!!!!!!!!!!!!! (id: ' + userId + ') was NOT added group Id ' + groupId + ')');
                    node.warn(error);
                    node.warn(error.message);
                });
            } catch (error) {
                node.warn(error);
				node.warn(error.message);
            }
        });
    }

    RED.nodes.registerType("azure_add_user_to_group", azure_add_user_to_group_Node);
};
