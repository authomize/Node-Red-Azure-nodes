const axios = require('axios');

module.exports = function(RED) {
    function azure_does_group_exist_Node(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        var results = [];
        
		node.auth = RED.nodes.getNode(config.auth);

        node.on('input', async function(msg) {
			if (!node.auth || !node.auth.has_credentials) {
				node.error("auth configuration is missing");
				return
			}

			const groupName = RED.util.evaluateNodeProperty(
				config.groupName, config.groupNameType, node, msg
			)

			const access_token = await node.auth.get_access_token();

            try {
				var searchURL = 'https://graph.microsoft.com/v1.0/groups/?$filter=displayName%20eq%20%27' + encodeURIComponent(groupName) + '%27';
				const response = await axios.get(searchURL, {
					headers: {
					  Authorization: 'Bearer ' + access_token
					}
				});

				if (response.data.value.length == 1) {
					msg.groupID = response.data.value[0].id;
					node.send([msg, null]);
				} else if (response.data.value.length > 1){
					throw new Error('There is more than one group with the suggested name');
				} else {
					node.send([null, msg])
				}
            } catch (error) {
                node.warn(error);
            }
        });
    }

    RED.nodes.registerType("azure_does_group_exist", azure_does_group_exist_Node);
};
