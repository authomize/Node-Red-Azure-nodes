const axios = require('axios');

module.exports = function(RED) {
    function azure_is_user_in_group_Node(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        var results = [];
        var access_token = '';
        
        node.on('input', async function(msg) {
            const access_token = msg.access_token;
			const groupId = msg.groupID;
			const groupName = msg.group_name;
			const user_id = msg.userIds[0];
			const user_name = msg.emails[0];
			
			
            try {
				const url = 'https://graph.microsoft.com/v1.0/users/' + user_id + '/memberOf';
				const headers = {Authorization: 'Bearer ' + access_token};
				const response = await axios.get(url, { headers });

				// Check if the group exists in the response
				const groups = response.data.value;
				const isMember = groups.some(group => group.id === groupId);

				if (isMember) {
					node.send([msg, null]);
				} else {
					throw new Error(user_name + ' (id: ' + user_id + ') is not a member of: ' + groupName);
					node.send([null, msg])
				}
            } catch (error) {
                node.warn(error);
            }
        });
    }

    RED.nodes.registerType("azure_is_user_in_group", azure_is_user_in_group_Node);
};