
        Dim urlMaillingSimplificado = IAgentToolbarHandler.AgentToolbarLinkSisWeb()

        'IRequestsHandler.SendRequest(urlMaillingSimplificado, "", "text/html; charset=utf-8", "GET", "")

        Dim returnUrl = urlMaillingSimplificado.ToString().Substring(urlMaillingSimplificado.ToString().IndexOf("Unificado"))

        IRequestsHandler.SendRequest(String.Format("http://operacao-sisweb/Unificado/TLV/Session/Login?returnURL={0}", returnUrl), "Matricula=TR739886&Senha=foc21are", "application/x-www-form-urlencoded", "POST", "")

        Dim returnURLLogado = urlMaillingSimplificado.ToString().Substring(urlMaillingSimplificado.ToString().IndexOf("fila_descricao"))

        IRequestsHandler.SendRequest(String.Format("http://operacao-sisweb/Unificado/TLV/InclusaoDeProtocolo/ViaUrl?{0}", returnURLLogado), "", "text/html; charset=utf-8", "GET", "")

