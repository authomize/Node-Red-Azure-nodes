const axios = require('axios');

module.exports = function(RED) {
    function azure_create_group(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        node.auth = RED.nodes.getNode(config.auth);
        
        node.on('input', async function(msg) {

            if (!node.auth || !node.auth.has_credentials) {
				node.error("auth configuration is missing");
				return
			}

			const access_token = await node.auth.get_access_token();

            const groupName = msg.group_name;
            const group_desc = msg.group_desc;
            const headers = {Authorization: 'Bearer ' + access_token, "Content-Type": 'application/json' };
            const url = 'https://graph.microsoft.com/v1.0/groups';

            const postData = {
                description: group_desc,
                displayName: groupName,
                groupTypes: [],
                mailEnabled: false,
                mailNickname: 'Authomize_' + groupName,
                securityEnabled: true
            };

			try {
				const response = await axios({
					method: 'post',
					url: url,
					headers: headers,
					data: postData
				});
				
				msg.groupID = response.data.id;
				node.send(msg);

			  } catch (error) {
					node.warn(error);
                    node.warn(error.message);
			  }

        });
    }
    RED.nodes.registerType("azure_create_group", azure_create_group);
};
