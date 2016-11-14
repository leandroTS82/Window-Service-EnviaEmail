using Microsoft.SharePoint.Client;
using Modelos;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Security;

namespace WindowServiceEnviaEmail
{
    public static class DataSharepoint
    {
        public static ClientContext AutenticaAcessoSPOnline(string url, string user, string password, out ErroModel registroErro)
        {
            ClientContext clientContex = new ClientContext(url);
            registroErro = null;
            try
            {
                var passWord = new SecureString();
                foreach (char c in password.ToCharArray()) passWord.AppendChar(c);
                clientContex.Credentials = new SharePointOnlineCredentials(user, passWord);
                clientContex.ExecuteQuery();
                return clientContex;
            }
            catch (Exception ex)
            {
                registroErro.Mensagem = $"O correu um erro durante a autenticação no sharepoint online no site {url}. {ex.Message}";
                registroErro.Detalhes = ex.StackTrace;
                registroErro.AppOrObjeto = ex.Source;
                return clientContex;
            }

        }
        public static List<ItensListaModel> obtemLista(ClientContext context, out ErroModel registroErro, string lista)
        {
            List<ItensListaModel> itens = new List<ItensListaModel>();
            registroErro = null;
            try
            {
                List spList = context.Web.Lists.GetByTitle(lista);
                CamlQuery camlQuery = new CamlQuery();
                //camlQuery.ViewXml = query;
                var listItemcollection = spList.GetItems(camlQuery);
                if (lista == "Teste")
                {
                    context.Load(
                        listItemcollection,
                                        items => items.Include(
                                        item => item["Title"],
                                        item => item["Pessoa"]
                                        ));
                    context.ExecuteQuery();
                    if (registroErro == null)
                        itens = TratamentoRetornoItens(listItemcollection, out registroErro);
                }
                else
                {
                    context.Load(
                           listItemcollection,
                                           items => items.Include(
                                           item => item["Title"]
                                           ));
                    context.ExecuteQuery();
                    if (registroErro == null)
                        itens = TratamentoRetornoItens(listItemcollection, out registroErro);
                }
            }
            catch (Exception ex)
            {
                registroErro.Mensagem = $"Ocorreu um erro durante a consulta da lista {lista}. Mensagem: {ex.Message}";
                registroErro.Detalhes = ex.StackTrace;
                registroErro.AppOrObjeto = ex.Source;
            }
            return itens;
        }
        private static List<ItensListaModel> TratamentoRetornoItens(ListItemCollection listItemcollection, out ErroModel registroErro)
        {
            List<ItensListaModel> itens = new List<ItensListaModel>();
            registroErro = null;
            try
            {
                foreach (var item in listItemcollection)
                {
                    // Obtem o ID do documento e verifica se o mesmo é um documento válido
                    ItensListaModel process = new ItensListaModel();
                    //verifica o tipo do campo obtido em cada item, verifica a propriedade e adiciona o valor correto.
                    foreach (var field in item.FieldValues)
                    {
                        if (field.Value is ClientValueObject)
                        {
                            try
                            {
                                string lookupValue = "";
                                int lookupId = 0;
                                FieldLookupValue lookupField = field.Value as FieldLookupValue;
                                if (lookupField != null)
                                {
                                    lookupValue = lookupField.LookupValue;
                                    lookupId = lookupField.LookupId;
                                }
                                else
                                {
                                    FieldUserValue lookupUser = field.Value as FieldUserValue;
                                    lookupValue = lookupUser.LookupValue;
                                    lookupId = lookupUser.LookupId;
                                }
                                PropertyInfo propId = process.GetType().GetProperty("ID_" + field.Key);
                                if (propId != null && propId.CanWrite) propId.SetValue(process, lookupId, null);
                                PropertyInfo propDesc = process.GetType().GetProperty("Nome_" + field.Key);
                                if (propDesc != null && propDesc.CanWrite) propDesc.SetValue(process, lookupValue, null);
                            }
                            catch (Exception ex)
                            {
                                registroErro.Mensagem = $"Ocorreu um erro durante o tratamento das propriedades para campos do tipo ClientValueObject. Mensagem: {ex.Message}";
                                registroErro.Detalhes = ex.StackTrace;
                                registroErro.AppOrObjeto = ex.Source;
                                continue;
                            }
                        }
                        else
                        {
                            try
                            {
                                PropertyInfo prop = process.GetType().GetProperty(field.Key);
                                if (prop != null && prop.CanWrite)
                                {
                                    var val = field.Value;
                                    if (prop.PropertyType.Name == "Boolean" || prop.PropertyType.Name == "Int32" || prop.PropertyType.Name == "Single") val = Convert.ToInt32(val);
                                    if (prop.PropertyType.Name == "DateTime") val = Convert.ToDateTime(val);
                                    if (prop.PropertyType.Name == "String" && val != null) val = val.ToString();
                                    prop.SetValue(process, val, null);
                                }
                            }
                            catch (Exception ex)
                            {
                                registroErro.Mensagem = $"Ocorreu um erro durante o tratamento dos campos genéricos. Mensagem: {ex.Message}";
                                registroErro.Detalhes = ex.StackTrace;
                                registroErro.AppOrObjeto = ex.Source;
                                continue;
                            }
                        }
                    }
                    itens.Add(process);
                }
            }
            catch (Exception ex)
            {
                registroErro.Mensagem = $"Ocorreu um erro durante o o tratamento de retorno dos itens. Mensagem: {ex.Message}";
                registroErro.Detalhes = ex.StackTrace;
                registroErro.AppOrObjeto = ex.Source;
            }
            return itens;
        }
    }
}
