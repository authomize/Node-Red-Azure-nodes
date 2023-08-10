const axios = require('axios');

module.exports = function (RED) {
  function azure_remove_user_from_app_Node(config) {
    RED.nodes.createNode(this, config);
    var node = this;

    node.auth = RED.nodes.getNode(config.auth);

    node.on('input', async function (msg) {
      if (!node.auth || !node.auth.has_credentials) {
				node.error("auth configuration is missing");
				return
			}

      const access_token = await node.auth.get_access_token();

      const headers = { Authorization: 'Bearer ' + access_token, 'Content-Type': 'application/json' };
      var app_object_ID = '';
      var principalId = '';
      var appRoleAssignedToId = '';
      var appName = '';
      var user_name = '';

      if (msg.payload.data.entities[0].object === 'asset') {
        app_object_ID = msg.payload.data.entities[0].originId;
        appName = msg.payload.data.entities[0].name;
        principalId = msg.payload.data.entities[1].originId;
        user_name = msg.payload.data.entities[1].name;
      } else {
        app_object_ID = msg.payload.data.entities[1].originId;
        appName = msg.payload.data.entities[1].name;
        principalId = msg.payload.data.entities[0].originId;
        user_name = msg.payload.data.entities[0].name;
      }

      var url = 'https://graph.microsoft.com/v1.0/servicePrincipals/' + app_object_ID + '/appRoleAssignedTo/';
      try {
        axios.get(url, { headers })
          .then(response => {
            const appRoleAssignments = response.data.value;
            for (const assignment of appRoleAssignments) {
              if (assignment.principalId === principalId) {
                appRoleAssignedToId = assignment.id;
                break;
              }
            }

            var url_delete = 'https://graph.microsoft.com/v1.0/servicePrincipals/' + app_object_ID + '/appRoleAssignedTo/' + appRoleAssignedToId;
            try {
              axios.delete(url_delete, { headers })
                .then(response => {
                  node.send(msg);
				  node.warn(user_name + ' (id: ' + principalId + ') was removed from ' + appName + ' (id: ' + app_object_ID + ')');
                })
                .catch(error => {
				  node.warn('!!!!!!!!!!!!!!!!!!!!' + user_name + ' (id: ' + principalId + ') was NOT removed from ' + appName + ' (id: ' + app_object_ID + ')');
                  node.warn(error);
                  node.warn(error.message);
                });
            } catch (error) {
              node.warn(error);
              node.warn(error.message);
            }
          })
          .catch(error => {
            node.warn(error);
            node.warn(error.message);
          });
      } catch (error) {
        node.warn(error);
        node.warn(error.message);
      }
    });
  }

  RED.nodes.registerType('azure_remove_user_from_app', azure_remove_user_from_app_Node);
};
