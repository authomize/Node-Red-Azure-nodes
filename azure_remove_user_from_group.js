const axios = require('axios');

module.exports = function(RED) {
    function azure_remove_user_from_group_Node(config) {
        RED.nodes.createNode(this, config);
        var node = this;
       
        node.on('input', async function(msg) {
            const access_token = msg.access_token;
            const groupID = msg.groupID;
            const userID = msg.usersId[0];
			const user_name = msg.emails[0];
			const groupName = msg.group_name;
			
			var url = 'https://graph.microsoft.com/v1.0/groups/' + groupID + '/members/' + userID + '/$ref';
            const headers = { Authorization: 'Bearer ' + access_token, "Content-Type": 'application/json' };
            
			try {
                axios.delete(url, { headers }).then(response => {
					node.warn(user_name + ' (id: ' + userID + ') was removed from ' + groupName + ' (id: ' + groupID + ')');
					node.send(msg);
                }).catch(error => {
                    node.warn(error);
                    node.warn(error.message);
                });
            } catch (error) {
                node.warn(error);
				node.warn(error.message);
            }
        });
    }

    RED.nodes.registerType("azure_remove_user_from_group", azure_remove_user_from_group_Node);
};
