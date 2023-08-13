const axios = require('axios');

module.exports = function(RED) {
    function azure_revoke_signin_sessions(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        node.auth = RED.nodes.getNode(config.auth);
        
        node.on('input', async function(msg) {

            if (!node.auth || !node.auth.has_credentials) {
				node.error("auth configuration is missing");
				return
			}

			const access_token = await node.auth.get_access_token();
			
			const userID = RED.util.evaluateNodeProperty(
				config.userID, config.userIDType, node, msg
			)
			const userName = RED.util.evaluateNodeProperty(
				config.userName, config.userNameType, node, msg
			)
			
			const headers = {Authorization: 'Bearer ' + access_token, "Content-Type": 'application/json' };
            const url = "https://graph.microsoft.com/v1.0/users/"  + userID + "/revokeSignInSessions";
			node.warn(url)
			try {
				const response = await axios({
					method: 'post',
					url: url,
					headers: headers
				});
				node.warn(userName + ' (id: ' + userID + ') Session was revoked');
				node.send(msg);

			  } catch (error) {
                    node.warn('!!!!!!!!!!!!!!!!!!!!' + userName + ' (id: ' + userID + ') Session was NOT revoked');
                    node.warn(error.message);
			  }

        });
    }
    RED.nodes.registerType("azure_revoke_signin_sessions", azure_revoke_signin_sessions);
};
