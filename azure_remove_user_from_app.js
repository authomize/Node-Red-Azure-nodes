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

      const appObjectId = RED.util.evaluateNodeProperty(
				config.appObjectId, config.appObjectIdType, node, msg
			  )

      const principalId = RED.util.evaluateNodeProperty(
				config.principalId, config.principalIdType, node, msg
			  )
		
			  const appRoleAssignedToId = RED.util.evaluateNodeProperty(
				config.appRoleAssignedToId, config.appRoleAssignedToIdType, node, msg
			  )

      var url = 'https://graph.microsoft.com/v1.0/servicePrincipals/' + appObjectId + '/appRoleAssignedTo/';
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

            var url_delete = 'https://graph.microsoft.com/v1.0/servicePrincipals/' + appObjectId + '/appRoleAssignedTo/' + appRoleAssignedToId;
            try {
              axios.delete(url_delete, { headers })
                .then(response => {
                  node.send(msg);
				          node.warn(principalId + ' was removed from ' + appObjectId);
                })
                .catch(error => {
				          node.warn('!!!!!!!!!!!!!!!!!!!!' + principalId + ' was NOT removed from ' + appObjectId);
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
