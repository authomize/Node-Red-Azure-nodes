const axios = require('axios');

module.exports = function(RED) {
    function azure_disable_user_Node(config) {
        RED.nodes.createNode(this, config);
        var node = this;

		node.auth = RED.nodes.getNode(config.auth);

        node.on('input', async function(msg) {
			if (!node.auth || !node.auth.has_credentials) {
				node.error("auth configuration is missing");
				return
			}

			const access_token = await node.auth.get_access_token();

			try {
				const userID = msg.usersId[0];
				const user_name = msg.emails[0];

				const url = 'https://graph.microsoft.com/v1.0/users/' + userID;
				const headers = { Authorization: 'Bearer ' + access_token, "Content-Type": 'application/json' };
				const data = {accountEnabled: false};

				const response = await axios.patch(url, data, { headers });
				node.warn(user_name + ' (id: ' + userID + ') was disabled');
				node.send(msg);
            } catch (error) {
                node.warn(error);
				node.warn(error.message);
            }
        });
    }

    RED.nodes.registerType("azure_disable_user", azure_disable_user_Node);
};
