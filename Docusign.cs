            string Username = ConfigurationManager.AppSettings.Get("Username");
            string Password = ConfigurationManager.AppSettings.Get("Password");
            string IntegratorKey = ConfigurationManager.AppSettings.Get("IntegratorKey"); 
            string basePath = ConfigurationManager.AppSettings.Get("basePath");
            if (optTipoVenda == null) return null;
            int alteracaoPosVenda = 0;
            int.TryParse(Request.QueryString["alteracao_posvenda"], out alteracaoPosVenda);
            int posvenada = 0;
            int.TryParse(Request.QueryString["posvenda"], out posvenada);

            ApiClient apiClient = new ApiClient(basePath);
            string authHeader = "{\"Username\":\"" + Username + "\", \"Password\":\"" + Password + "\", \"IntegratorKey\":\"" + IntegratorKey + "\"}";
            apiClient.Configuration.AddDefaultHeader("X-DocuSign-Authentication", authHeader);

            // we will retrieve this from the login() results
            string accountId = null;

            // the authentication api uses the apiClient (and X-DocuSign-Authentication header) that are set in Configuration object
            AuthenticationApi authApi = new AuthenticationApi(apiClient);
            LoginInformation loginInfo = authApi.Login();

            // user might be a member of multiple accounts
            accountId = loginInfo.LoginAccounts[0].AccountId;
