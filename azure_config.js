const axios = require("axios");

module.exports = function (RED) {
  function AzureAdConfigNode(config) {
    RED.nodes.createNode(this, config);

    // Validate if the required JSON field is present
    if (!this.credentials || !this.credentials.tenantId) {
      this.error("Tenant Id is missing");
      return;
    }

    if (!this.credentials || !this.credentials.clientId) {
      this.error("Client id is missing");
      return;
    }

    if (!this.credentials || !this.credentials.clientSecret) {
      this.error("Client secret is missing");
      return;
    }

    this.name = config.name;
    this.tenantId = this.credentials.tenantId;
    this.clientId = this.credentials.clientId;
    this.clientSecret = this.credentials.clientSecret;

    this.has_credentials = this.tenantId && this.clientId && this.clientSecret;

    async function generateToken() {
      const formUrlEncoded = (x) =>
        Object.keys(x).reduce(
          (p, c) => p + `&${c}=${encodeURIComponent(x[c])}`,
          ""
        );

      try {
        const clientId = this.clientId;
        const clientSecret = this.clientSecret;
        const tenantId = this.tenantId;

        const tokenEndpoint =
          "https://login.microsoftonline.com/" +
          tenantId +
          "/oauth2/v2.0/token";
        const bodyFormData = {
          grant_type: "client_credentials",
          client_id: clientId,
          client_secret: clientSecret,
          scope: "https://graph.microsoft.com/.default",
        };

        const response = await axios({
          url: tokenEndpoint,
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          data: formUrlEncoded(bodyFormData),
        });

        return response.data.access_token;
      } catch (error) {
        console.error(error);
      }
    };

    this.get_access_token = generateToken;
  }

  RED.nodes.registerType("azure_config", AzureAdConfigNode, {
    credentials: {
      tenantId: { type: "text" },
      clientId: { type: "text" },
      clientSecret: { type: "text" },
    },
  });
};
