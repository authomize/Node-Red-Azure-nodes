const axios = require('axios');

module.exports = function(RED) {
    function azure_is_user_in_group_Node(config) {
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
				const url = `https://graph.microsoft.com/v1.0/users/${userId}/memberOf`;
				const headers = {Authorization: 'Bearer ' + access_token};
				const response = await axios.get(url, { headers });

				// Check if the group exists in the response
				const groups = response.data.value;
				const isMember = groups.some(group => group.id === groupId);

				if (isMember) {
					node.send([msg, null]);
				} else {
					throw new Error(user_id + ' is not a member of: ' + groupId);
					node.send([null, msg])
				}
            } catch (error) {
                node.warn(error);
            }
        });
    }

    RED.nodes.registerType("azure_is_user_in_group", azure_is_user_in_group_Node);
};
