const axios = require('axios');

module.exports = function(RED) {
    function azure_add_user_to_group_Node(config) {
        RED.nodes.createNode(this, config);
        var node = this;
       
        node.on('input', async function(msg) {
			try {
				const access_token = msg.access_token;
				const groupID = msg.groupID;
				const userID = msg.usersId[0];
				const user_name = msg.emails[0];
				const groupName = msg.group_name;
				
				var url = 'https://graph.microsoft.com/v1.0/groups/' + groupID + '/members/$ref';
				
				const postData = {
					"@odata.id": 'https://graph.microsoft.com/v1.0/directoryObjects/' + userID
				};
				
				const headers = { Authorization: 'Bearer ' + access_token, "Content-Type": 'application/json' };
            
                axios.post(url, postData, { headers }).then(response => {
					node.warn(user_name + ' (id: ' + userID + ') was added to ' + groupName + ' (id: ' + groupID + ')');
					node.send(msg);
                }).catch(error => {
					node.warn('!!!!!!!!!!!!!!!!!!!!' + user_name + ' (id: ' + userID + ') was NOT added to ' + groupName + ' (id: ' + groupID + ')');
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
