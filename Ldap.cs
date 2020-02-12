using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Ports;
using System.Text;

/* used for ad */
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;

using ScriptStack;
using ScriptStack.Compiler;
using ScriptStack.Runtime;

namespace github.com.mbrain
{

    public class Ldap : Stack
    {

        private static ReadOnlyCollection<CustomMethod> prototypes;

        public ReadOnlyCollection<CustomMethod> Prototypes
        {
            get { return prototypes; }
        }
        public Ldap() {

            if (prototypes != null) return;
            List<CustomMethod> customMethods = new List<CustomMethod>
            {
                new CustomMethod(typeof(ArrayList), "getLocalUsers", "Get a list of all local users."),
                new CustomMethod((Type)null, "addLocalUser", typeof(string), typeof(string), "Add a local system user."),
                new CustomMethod((Type)null, "removeLocalUser", typeof(string), "Remove a local system user."),
                new CustomMethod(typeof(ArrayList), "getLocalGroupMembers", typeof(string), "Get all members of a local user group."),
                new CustomMethod(typeof(ArrayList), "getDomains", "Get all domains in the current forrest."),
                new CustomMethod(typeof(ArrayList), "getGlobalCatalogs", "Get all global catalogs in the current forrest."),
                new CustomMethod(typeof(ArrayList), "getDomainControllers", "Get all domain controllers in the current forrest."),
                new CustomMethod((Type)null, "createTrustRelationship", (Type)null, "Create a trust relationship between 2 forrests."),
                new CustomMethod((Type)null, "deleteTrustRelationship", (Type)null, "Delete a trust relationship between 2 forrests."),
                new CustomMethod(typeof(ArrayList), "objectsInDn", (Type)null, "Get all objects in a specific Dn."),
                /* todo
                new CustomMethod(typeof(bool), "directoryEntryConfigurationSettings", (Type)null, "Get the config settings of an entry."),
                */
                new CustomMethod(typeof(bool), "objectExists", (Type)null, "Check if an object exists by its fqdn."),
                new CustomMethod(typeof(bool), "moveObject", (Type)null, "Move an object using its fqdn."),
                /* todo
                new CustomMethod(typeof(bool), "attributeValuesMultiString", (Type)null, "TODO"),
                new CustomMethod(typeof(bool), "attributeValuesSingleString", (Type)null, "TODO"),
                */
                new CustomMethod(typeof(ArrayList), "getUsedAttributes", (Type)null, "Get all attributes an objects uses by its fqdn."),
                new CustomMethod(typeof(bool), "addUserToADGroup", (Type)null, "Add a user to a group."),
                new CustomMethod(typeof(bool), "removeUserFromADGroup", (Type)null, "Remove a user from a group.")
                /* todo *
                new CustomMethod(typeof(bool), "getUserGroupMemberships", (Type)null, "Get all groups a user is member of.")
                */
            };
            prototypes = customMethods.AsReadOnly();
        }

        public object OnMethodInvoke(String functionName, List<object> parameters) {

            /* add local user */
            if(functionName == "addLocalUser")
            {
                try
                {
                    DirectoryEntry localMachine = new DirectoryEntry("WinNT://" + Environment.MachineName);
                    DirectoryEntry user = localMachine.Children.Add((string)parameters[0], "user");
                    user.Invoke("SetPassword", new object[] { (string)parameters[1] });
                    user.CommitChanges();
                    string result = user.Guid.ToString();
                    localMachine.Close();
                    user.Close();
                    return result;
                }catch(Exception) { return "-1"; }
            }

            /* delete local user */
            if(functionName == "removeLocalUser")
            {
                try { 
                    DirectoryEntry localMachine = new DirectoryEntry("WinNT://" + Environment.MachineName);
                    DirectoryEntry user = localMachine.Children.Find((string)parameters[0], "user");
                    localMachine.Children.Remove(user);
                    //user.CommitChanges();
                    localMachine.Close();
                    user.Close();
                }catch (Exception) { return false; }
                return true;
            }

            if(functionName == "getLocalUsers")
            {
                DirectoryEntry localMachine = new DirectoryEntry("WinNT://" + Environment.MachineName + ", Computer");
                ArrayList users = new ArrayList();
                foreach (DirectoryEntry child in localMachine.Children)
                {
                    if (child.SchemaClassName == "User")
                        users.Add(child.Name);
                }
                return users;
            }

            if(functionName == "getLocalGroupMembers")
            {
                DirectoryEntry localMachine = new DirectoryEntry("WinNT://" + Environment.MachineName + ", Computer");
                DirectoryEntry group = localMachine.Children.Find((string)parameters[0], "group");
                object members = group.Invoke("members", null);
                ArrayList all = new ArrayList();
                foreach (object groupMember in (System.Collections.IEnumerable)members)
                {
                    DirectoryEntry member = new DirectoryEntry(groupMember);
                    all.Add(member.Name);
                }
                return all;
            }

            if (functionName == "getDomains")
            {
                ArrayList allDomains = new ArrayList();
                Forest currentForest = Forest.GetCurrentForest();
                DomainCollection myDomains = currentForest.Domains;

                foreach (Domain objDomain in myDomains)
                {
                    allDomains.Add(objDomain.Name);
                }
                return allDomains;
            }
            if (functionName == "getGlobalCatalogs")
            {
                ArrayList alGCs = new ArrayList();
                Forest currentForest = Forest.GetCurrentForest();
                foreach (GlobalCatalog gc in currentForest.GlobalCatalogs)
                {
                    alGCs.Add(gc.Name);
                }
                return alGCs;
            }
            if (functionName == "getDomainControllers")
            {
                ArrayList alDcs = new ArrayList();
                Domain domain = Domain.GetCurrentDomain();
                foreach (DomainController dc in domain.DomainControllers)
                {
                    alDcs.Add(dc.Name);
                }
                return alDcs;
            }
            if (functionName == "createTrustRelashionship")
            {
                Forest sourceForest = Forest.GetForest(new DirectoryContext(DirectoryContextType.Forest, (string)parameters[0]));
                Forest targetForest = Forest.GetForest(new DirectoryContext(DirectoryContextType.Forest, (string)parameters[1]));
                sourceForest.CreateTrustRelationship(targetForest,TrustDirection.Outbound);
            }
            if (functionName == "deleteTrustRelashionship")
            {
                Forest sourceForest = Forest.GetForest(new DirectoryContext(DirectoryContextType.Forest, (string)parameters[0]));
                Forest targetForest = Forest.GetForest(new DirectoryContext(DirectoryContextType.Forest, (string)parameters[1]));
                sourceForest.DeleteTrustRelationship(targetForest);
            }
            if (functionName == "objectsInDn")
            {
                ArrayList allObjects = new ArrayList();
                try
                {
                    DirectoryEntry directoryObject = new DirectoryEntry("LDAP://" + (string)parameters[0]);
                    foreach (DirectoryEntry child in directoryObject.Children)
                    {
                        string childPath = child.Path.ToString();
                        allObjects.Add(childPath.Remove(0, 7));
                        //remove the LDAP prefix from the path
                        child.Close();
                        child.Dispose();
                    }
                    directoryObject.Close();
                    directoryObject.Dispose();
                }
                catch (DirectoryServicesCOMException) {}
                return allObjects;
            }
            if (functionName == "directoryEntryConfigurationSettings")
            {
                // Bind to current domain
                DirectoryEntry entry = new DirectoryEntry((string)parameters[0]);
                DirectoryEntryConfiguration entryConfiguration = entry.Options;
                // todo
                /*
                Console.WriteLine("Server: " + entryConfiguration.GetCurrentServerName());
                Console.WriteLine("Page Size: " + entryConfiguration.PageSize.ToString());
                Console.WriteLine("Password Encoding: " + entryConfiguration.PasswordEncoding.ToString());
                Console.WriteLine("Password Port: " + entryConfiguration.PasswordPort.ToString());
                Console.WriteLine("Referral: " + entryConfiguration.Referral.ToString());
                Console.WriteLine("Security Masks: " + entryConfiguration.SecurityMasks.ToString());
                Console.WriteLine("Is Mutually Authenticated: " + entryConfiguration.IsMutuallyAuthenticated().ToString());
                Console.WriteLine();
                Console.ReadLine();
                */
            }
            if (functionName == "objectExists")
            {
                return (DirectoryEntry.Exists("LDAP://" + (string)parameters[0]));
            }
            if (functionName == "moveObject")
            {
                DirectoryEntry eLocation = new DirectoryEntry("LDAP://" + (string)parameters[0]);
                DirectoryEntry nLocation = new DirectoryEntry("LDAP://" + (string)parameters[1]);
                string newName = eLocation.Name;
                eLocation.MoveTo(nLocation, newName);
                nLocation.Close();
                eLocation.Close();
            }
            if (functionName == "attributeValuesMultiString")
            {
                string objectDn = (string)parameters[0];
                string attributeName = (string)parameters[1];
                ArrayList valuesCollection = (ArrayList)parameters[2];

                DirectoryEntry ent = new DirectoryEntry(objectDn);
                PropertyValueCollection ValueCollection = ent.Properties[attributeName];
                System.Collections.IEnumerator en = ValueCollection.GetEnumerator();

                while (en.MoveNext())
                {
                    if (en.Current != null)
                    {
                        if (!valuesCollection.ToString().Contains(en.Current.ToString()))
                        {
                            valuesCollection.Add(en.Current.ToString());
                            //if (recursive) -> do it again
                        }
                    }
                }
                ent.Close();
                ent.Dispose();
                return valuesCollection;
            }
            if (functionName == "attributeValuesSingleString")
            {
                string objectDn = (string)parameters[0];
                string attributeName = (string)parameters[1];
                string strValue;
                DirectoryEntry ent = new DirectoryEntry(objectDn);
                strValue = ent.Properties[attributeName].Value.ToString();
                ent.Close();
                ent.Dispose();
                return strValue;
            }
            if (functionName == "getUsedAttributes")
            {
                DirectoryEntry obj = new DirectoryEntry("LDAP://" + (string)parameters[0]);
                ArrayList props = new ArrayList();
                foreach (string strAttrName in obj.Properties.PropertyNames)
                {
                    props.Add(strAttrName);
                }
                return props;
            }
            if (functionName == "addUserToADGroup")
            {
                string groupDn = (string)parameters[0];
                string userDn = (string)parameters[1];
                try
                {
                    DirectoryEntry dirEntry = new DirectoryEntry("LDAP://" + groupDn);
                    dirEntry.Properties["member"].Add(userDn);
                    dirEntry.CommitChanges();
                    dirEntry.Close();
                }
                catch(DirectoryServicesCOMException) {}
            }
            if (functionName == "removeUserFromADGroup")
            {
                string groupDn = (string)parameters[0];
                string userDn = (string)parameters[1];
                try
                {
                    DirectoryEntry dirEntry = new DirectoryEntry("LDAP://" + groupDn);
                    dirEntry.Properties["member"].Remove(userDn);
                    dirEntry.CommitChanges();
                    dirEntry.Close();
                }
                catch(DirectoryServicesCOMException) {}
            }
            if (functionName == "getUserGroupMemberships")
            {
                string groupDn = (string)parameters[0];
                string userDn = (string)parameters[1];
                ArrayList groupMemberships = new ArrayList();
                // todo
                //return AttributeValuesMultiString("memberOf", userDn, groupMemberships, recursive);
            }

            /* todo: enumerate a DN 
             * see: https://stackoverflow.com/questions/26744000/get-email-addresses-from-exchange-given-directoryentry
             */
            if (functionName == "enum")
            {
                string ConnectionString = String.Empty;
                string ProviderUserName = String.Empty;
                string ProviderPassword = String.Empty;
                string username = String.Empty;
                DirectoryEntry directoryEntry = new DirectoryEntry(ConnectionString, ProviderUserName, ProviderPassword, AuthenticationTypes.Secure);

                DirectorySearcher search = new DirectorySearcher(directoryEntry);
                search.Filter = "(&(objectClass=user)(sAMAccountName=" + username + "))";
                search.CacheResults = false;

                SearchResultCollection allResults = search.FindAll();
                StringBuilder sb = new StringBuilder();

                foreach (SearchResult searchResult in allResults)
                {
                    foreach (string propName in searchResult.Properties.PropertyNames)
                    {
                        ResultPropertyValueCollection valueCollection = searchResult.Properties[propName];
                        foreach (Object propertyValue in valueCollection)
                        {
                            sb.AppendLine(string.Format("property:{0}, value{1}", propName, propertyValue));
                        }
                    }
                }

                return sb.ToString();
            }
            return null;
        }

    }

}
