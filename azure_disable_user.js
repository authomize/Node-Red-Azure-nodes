const axios = require('axios');

module.exports = function(RED) {
    function azure_disable_user_Node(config) {
        RED.nodes.createNode(this, config);
        var node = this;

        node.on('input', async function(msg) {
			try {
				const access_token = msg.access_token;
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
