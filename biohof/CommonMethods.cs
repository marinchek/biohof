using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace R.iT.UtilityClasses
{
    public static class CommonMethods
    {
        /// <summary>
        /// Zusammenführung der Target Entity und des PreImage
        /// </summary>
        public static Entity AssemblePostImage(Entity targetEntity, Entity preImage)
        {
            Entity entity = new Entity(targetEntity.LogicalName);
            entity.Id = targetEntity.Id;

            foreach (var attribute in preImage.Attributes)
            {

                if (!targetEntity.Attributes.Contains(attribute.Key))
                {
                    entity.Attributes.Add(attribute);
                }
                else
                {
                    entity.Attributes.Add(targetEntity.Attributes.First(x => x.Key == attribute.Key));
                }
            }

            foreach (var attribute in targetEntity.Attributes)
            {
                if (!entity.Attributes.Contains(attribute.Key))
                {
                    entity.Attributes.Add(attribute);
                }
            }

            return entity;
        }

        private static readonly byte[] Va = { 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x11, 0x11, 0x12, 0x13, 0x14, 0x0e, 0x16, 0x17 };

        public static string Decrypt(string text, string salt)
        {
            try
            {
                using (Aes aes = new AesManaged())
                {
                    Rfc2898DeriveBytes deriveBytes = new Rfc2898DeriveBytes(Encoding.UTF8.GetString(Va, 0, Va.Length), Encoding.UTF8.GetBytes(salt));
                    aes.Key = deriveBytes.GetBytes(128 / 8);
                    aes.IV = aes.Key;

                    using (MemoryStream decryptionStream = new MemoryStream())
                    {
                        using (CryptoStream decrypt = new CryptoStream(decryptionStream, aes.CreateDecryptor(), CryptoStreamMode.Write))
                        {
                            byte[] encryptedData = Convert.FromBase64String(text);

                            decrypt.Write(encryptedData, 0, encryptedData.Length);
                            decrypt.Flush();
                        }

                        byte[] decryptedData = decryptionStream.ToArray();
                        string decryptedText = Encoding.UTF8.GetString(decryptedData, 0, decryptedData.Length);

                        return decryptedText;
                    }
                }
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Gibt erstes Objekt zurück (aktiv / inaktiv)
        /// </summary>
        public static Entity RetrieveFirstByObject(IOrganizationService service, string entitySchemaname, string searchAttributeSchemaname, object searchObject, string[] set = null)
        {
            QueryExpression query = new QueryExpression();

            query.EntityName = entitySchemaname;
            query.ColumnSet = set == null ? new ColumnSet(true) : new ColumnSet(set);

            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.And;

            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = searchAttributeSchemaname;
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(searchObject);


            query.Criteria.Conditions.Add(condition1);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count != 0)
            {
                return results[0];
            }
            return null;
        }

        /// <summary>
        /// Gibt alle relevanten Objekte zurück (aktiv / inaktiv). Liefert maximal 5000 Datensätze zurück.
        /// </summary>
        public static EntityCollection RetrieveByObject(IOrganizationService service, string entitySchemaname, string searchAttributeSchemaname, object searchObject, string[] set = null)
        {
            QueryExpression query = new QueryExpression();

            query.EntityName = entitySchemaname;
            query.ColumnSet = set == null ? new ColumnSet(true) : new ColumnSet(set);

            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.And;

            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = searchAttributeSchemaname;
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(searchObject);

            query.Criteria.Conditions.Add(condition1);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count != 0)
            {
                return results;
            }
            return null;
        }

        public static Entity GetRpWebConfig(IOrganizationService service)
        {
            QueryExpression query = new QueryExpression();

            query.EntityName = "rit_rpwebkonfiguration";
            query.ColumnSet = new ColumnSet(true);

            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.And;

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count != 0)
            {
                return results.Entities.First();
            }
            return null;
        }

        public static EntityCollection RetrieveRealtedRecords(IOrganizationService service, string entitySchemaname1, string entitySchemaname2, string primaryfieldEntity1, string primaryfieldEntity2, string searchAttributeSchemanameEntity1, string searchAttributeSchemanameEntity2, string relationshipentityName)
        {
            QueryExpression query = new QueryExpression(entitySchemaname1);

            query.ColumnSet = new ColumnSet(true);

            LinkEntity linkEntity1 = new LinkEntity(entitySchemaname1, relationshipentityName, searchAttributeSchemanameEntity1, primaryfieldEntity1, JoinOperator.Inner);

            LinkEntity linkEntity2 = new LinkEntity(relationshipentityName, entitySchemaname2, searchAttributeSchemanameEntity2, primaryfieldEntity2, JoinOperator.Inner);

            linkEntity1.LinkEntities.Add(linkEntity2);

            query.LinkEntities.Add(linkEntity1);

            EntityCollection collRecords = service.RetrieveMultiple(query);
            return collRecords;
        }

        /// <summary>
        /// Prüft, ob im CRM bereits eine Stammnotiz am aktuellen Tag zu der genannten Adresse erstellt wurde
        /// </summary>
        public static bool CheckForExistingEmbargoStammnotizOnCurrentDay(IOrganizationService service, Guid accountId, Guid embargoStammnotizSchluesselId)
        {

            QueryExpression query = new QueryExpression();

            query.EntityName = "rit_stammnotiz";
            query.ColumnSet = new ColumnSet("rit_stammnotizid");

            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.And;

            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "rit_firmaid_lu";
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(accountId);

            query.Criteria.Conditions.Add(condition1);

            ConditionExpression condition2 = new ConditionExpression();
            condition2.AttributeName = "createdon";
            condition2.Operator = ConditionOperator.Today;

            query.Criteria.Conditions.Add(condition2);

            ConditionExpression condition3 = new ConditionExpression();
            condition3.AttributeName = "rit_stammnotizschluesselid_lu";
            condition3.Operator = ConditionOperator.Equal;
            condition3.Values.Add(embargoStammnotizSchluesselId);

            query.Criteria.Conditions.Add(condition3);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            EntityCollection results = service.RetrieveMultiple(query);

            //Ist keine Stammnotiz am heutigen Tag erstellt worden wird false zurückgegeben
            return results.Entities.Count != 0;
        }

        public static EntityCollection GetAssociatedEntityItems(IOrganizationService service, string relationshipName, string relatedEntityName, string entityName, Guid entityId, string[] columns = null)
        {
            EntityCollection result = null;
            QueryExpression query = new QueryExpression();
            query.EntityName = relatedEntityName;
            query.ColumnSet = new ColumnSet(columns);
            Relationship relationship = new Relationship();
            relationship.SchemaName = relationshipName;
            relationship.PrimaryEntityRole = EntityRole.Referencing;
            RelationshipQueryCollection relatedEntity = new RelationshipQueryCollection();
            relatedEntity.Add(relationship, query);
            RetrieveRequest request = new RetrieveRequest();
            request.RelatedEntitiesQuery = relatedEntity;
            request.ColumnSet = new ColumnSet(true);

            request.Target = new EntityReference
            {
                Id = entityId,
                LogicalName = entityName
            };

            RetrieveResponse response = (RetrieveResponse)service.Execute(request);
            RelatedEntityCollection relatedEntityCollection = response.Entity.RelatedEntities;
            if (relatedEntityCollection.Count > 0 && relatedEntityCollection.Values.Count > 0)
            {
                result = relatedEntityCollection.Values.ElementAt(0);
            }
            return result;
        }

        /// <summary>
        /// Kontrolliert, ob bereits eine M:N-Beziehung zwischen den angegebenen Elementen angelegt wurde
        /// </summary>
        public static bool ConfirmExistingManyToManyRelationshipRecord(IOrganizationService service, string relationshipName, string primaryEntityIdField, Guid primaryEntityId, string secondaryEntityIdField, Guid secondaryEntityId)
        {
            QueryExpression query = new QueryExpression();

            query.EntityName = relationshipName;
            query.ColumnSet = new ColumnSet(true);
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.And;

            //Leasinganfrage
            ConditionExpression conditionLa = new ConditionExpression();
            conditionLa.AttributeName = primaryEntityIdField;
            conditionLa.Operator = ConditionOperator.Equal;
            conditionLa.Values.Add(primaryEntityId);
            query.Criteria.Conditions.Add(conditionLa);

            //Genehmigungsformular
            ConditionExpression conditionGenF = new ConditionExpression();
            conditionGenF.AttributeName = secondaryEntityIdField;
            conditionGenF.Operator = ConditionOperator.Equal;
            conditionGenF.Values.Add(secondaryEntityId);
            query.Criteria.Conditions.Add(conditionGenF);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            EntityCollection results = service.RetrieveMultiple(query);

            //Gebe false zurück wenn Entities null ist oder 0 Datensätze enthält
            return results.Entities?.Count > 0;
        }

        /// <summary>
        /// Verbindet zwei Datensätze (m:n) miteinander
        /// </summary>
        public static void DoAssociateRequest(IOrganizationService service, EntityReference targetEntityReference, EntityReferenceCollection relatedEntityReferences, string referenceName)
        {
            AssociateRequest associateRequest = new AssociateRequest
            {
                Target = new EntityReference(targetEntityReference.LogicalName, targetEntityReference.Id),
                RelatedEntities = relatedEntityReferences,
                Relationship = new Relationship(referenceName)
            };

            service.Execute(associateRequest);
        }

        /// <summary>
        /// Gibt Datensätze anhand QueryExpression zurück (inkl. paging)
        /// </summary>
        /// <param name="service">CRM Proxy</param>
        /// <param name="query">Abfrage</param>
        /// <param name="pages">Limitiert die Maximalanzahl der zurückgegeben Pages. 0 = alle Pages</param>
        /// <param name="pageSize">Anzahl an Datensätze innerhalb einer Page</param>
        public static EntityCollection RetrieveByQueryExpression(IOrganizationService service, QueryExpression query, int pages = 0, int pageSize = 0)
        {
            EntityCollection returnCollection = new EntityCollection();

            int fetchCount = pageSize;
            if (pageSize == 0) fetchCount = 5000;

            const int pageNumber = 1;

            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = fetchCount;
            query.PageInfo.PageNumber = pageNumber;

            query.PageInfo.PagingCookie = null;

            EntityCollection results = service.RetrieveMultiple(query);
            int currentPage = 1;

            if (results.Entities != null && results.Entities.Count != 0)
            {
                returnCollection.Entities.AddRange(results.Entities);

                //Wenn pages = 0 ist gibt es keine Limitierung der Pages
                while (results.MoreRecords && (pages == 0 || currentPage < pages))
                {
                    currentPage++;
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = results.PagingCookie;

                    results = service.RetrieveMultiple(query);
                    returnCollection.Entities.AddRange(results.Entities);
                }
            }

            return returnCollection;
        }

        /// <summary>
        /// Gibt alle relevanten Objekte zurück (nur aktive Datensätze)
        /// </summary>
        /// <param name="service"></param>
        /// <param name="entitySchemaname"></param>
        /// <param name="searchAttributeSchemaname"></param>
        /// <param name="searchObject"></param>
        /// <param name="set">Spalten der Rückgabe</param>
        /// <returns></returns>
        public static EntityCollection RetrieveActiveByObject(IOrganizationService service, string entitySchemaname, string searchAttributeSchemaname, object searchObject, string[] set = null)
        {
            QueryExpression query = new QueryExpression();

            query.EntityName = entitySchemaname;

            query.ColumnSet = set == null ? new ColumnSet(true) : new ColumnSet(set);

            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.And;

            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = searchAttributeSchemaname;
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(searchObject);

            ConditionExpression condition2 = new ConditionExpression();
            condition2.AttributeName = "statecode";
            condition2.Values.Add(0);
            condition2.Operator = ConditionOperator.Equal;

            query.Criteria.Conditions.Add(condition1);
            query.Criteria.Conditions.Add(condition2);

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count != 0)
            {
                return results;
            }
            return null;
        }

        /// <summary>
        /// Gibt Datensätze anhand FetchExpression zurück
        /// </summary>
        public static EntityCollection RetrieveByFetchExpression(IOrganizationService service, FetchExpression query, int pages = 0, int pageSize = 0)
        {
            EntityCollection returnCollection = new EntityCollection();
            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities != null && results.Entities.Count != 0)
            {
                return returnCollection;
            }

            return returnCollection;
        }

        /// <summary>
        /// Führt einen SetState Request durch
        /// Dient der Aktivierung / Deaktivierung von Datensätzen
        /// </summary>
        public static void SetStateRequest(IOrganizationService service, string entityName, Guid entityId, int state, int status)
        {
            SetStateRequest setStateRequest = new SetStateRequest
            {
                EntityMoniker = new EntityReference(entityName, entityId),
                State = new OptionSetValue(state),
                Status = new OptionSetValue(status)
            };
            service.Execute(setStateRequest);
        }

        /// <summary>
        /// Statusänderung eines Opportunity-Datensatz auf gewonnen
        /// </summary>
        public static void SetOpportunityStateWin(IOrganizationService service, EntityReference recordRef, int statusInt)
        {
            WinOpportunityRequest request = new WinOpportunityRequest();
            Entity opportunityClose = new Entity("opportunityclose");
            opportunityClose.Attributes.Add("opportunityid", new EntityReference("opportunity", recordRef.Id));
            request.OpportunityClose = opportunityClose;
            OptionSetValue newStatus = new OptionSetValue();
            newStatus.Value = statusInt;
            request.Status = newStatus;
            service.Execute(request);
        }

        /// <summary>
        /// Statusänderung eines Opportunity-Datensatz auf verloren
        /// </summary>
        public static void SetOpportunityStateLose(IOrganizationService service, EntityReference recordRef, int statusInt)
        {
            LoseOpportunityRequest request = new LoseOpportunityRequest();
            Entity opportunityClose = new Entity("opportunityclose");
            opportunityClose.Attributes.Add("opportunityid", new EntityReference("opportunity", recordRef.Id));
            request.OpportunityClose = opportunityClose;
            OptionSetValue newStatus = new OptionSetValue();
            newStatus.Value = statusInt;
            request.Status = newStatus;
            service.Execute(request);
        }

        /// <summary>
        /// Abruf eines Leasinganfragen Teams (Teamstruktur-Datensätze) mit einer speziellen Funktion (bspw. Vertrieb) zu einer angegebenen Leasinganfrage
        /// </summary>
        public static Entity RetrieveLeasinganfragenTeamForOpportunityByFunction(IOrganizationService service, Guid opportunityId, string funktion)
        {
            QueryExpression queryExpression = new QueryExpression("rit_leasinganfragenteam");
            queryExpression.ColumnSet = new ColumnSet(true);
            queryExpression.Criteria.AddCondition("rit_leasinganfrageid_lu", ConditionOperator.Equal, opportunityId);
            queryExpression.Criteria.AddCondition("rit_rpffunktionid_luname", ConditionOperator.Equal, funktion);
            queryExpression.Criteria.FilterOperator = LogicalOperator.And;
            EntityCollection leasinganfragenTeamCollection = RetrieveByQueryExpression(service, queryExpression, 1, 5000);
            if (leasinganfragenTeamCollection == null || leasinganfragenTeamCollection.Entities == null || leasinganfragenTeamCollection.Entities.Count == 0) return null;
            return leasinganfragenTeamCollection.Entities.FirstOrDefault();
        }
    }

    public enum DataOperation
    {
        Create = 0,
        Update = 1,
        Delete = 2
    }
}
