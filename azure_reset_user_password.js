const axios = require('axios');

module.exports = function(RED) {
    function azure_reset_user_password_Node(config) {
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

				const url = 'https://graph.microsoft.com/v1.0/users/' + userID + '/authentication/methods/28c10230-6103-485e-b985-444c60001490/resetPassword';
				node.warn(url);
				const headers = { Authorization: 'Bearer ' + access_token, "Content-Type": 'application/json' };
				const data = {
				"requireChangeOnNextSignIn": true,
				"newPassword": "LoKipud280("
				};

				const response = await axios.post(url, data, { headers });
				node.warn(user_name + ' (id: ' + userID + ') password reset');
				node.send(msg);
            } catch (error) {
                node.warn(error);
				node.warn(error.message);
            }
        });
    }

    RED.nodes.registerType("azure_reset_user_password", azure_reset_user_password_Node);
};
