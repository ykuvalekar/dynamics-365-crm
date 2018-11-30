using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace Dynamics.CRM.YM.Plugins
{
    public class GetFieldSecurityProfileAssociationAction : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            string output = string.Empty;
            try
            {
                tracer.Trace("Execution started.");
                context.OutputParameters["output"] = output;
                if (context.InputParameters.Contains("fieldsecurityprofilename") &&
                    context.InputParameters.Contains("primaryentity") &&
                    context.InputParameters.Contains("entityid")
                    )
                {
                    tracer.Trace("Found required input parameters.");
                    switch ((string)context.InputParameters["primaryentity"])
                    {
                        case "systemuser":
                            tracer.Trace("Inside system user case.");
                            SystemUserProfile sup = CheckUserProfileAssociation(context, service, tracer);
                            context.OutputParameters["output"] = GenerateOutputString(sup.dctFieldPermissions);
                            break;
                        case "team":
                            tracer.Trace("Inside team case.");
                            TeamProfile tp = CheckTeamProfileAssociation(context, service, tracer);
                            context.OutputParameters["output"] = GenerateOutputString(tp.dctFieldPermissions);
                            break;
                        default:
                            tracer.Trace("Invalid case.");
                            break;
                    }
                }
                tracer.Trace("Execution ended.");
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        private string GenerateOutputString(Dictionary<string, FieldPermissions> dctOutput)
        {
            string output = string.Empty;
            if (dctOutput.Count > 0)
            {
                foreach (KeyValuePair<string, FieldPermissions> item in dctOutput)
                {
                    if (!string.IsNullOrEmpty(output))
                    {
                        output += ",";
                    }
                    output += "{'logicalName':'" + item.Key + "', 'canCreate':" + item.Value.canCreate + ", 'canUpdate':" + item.Value.canUpdate + ", 'canRead':" + item.Value.canRead + "}";
                }
            }
            return "[" + output + "]";
        }

        public SystemUserProfile CheckUserProfileAssociation(IPluginExecutionContext context, IOrganizationService service, ITracingService tracer)
        {
            SystemUserProfile sup = new SystemUserProfile();
            string securityProfileName = (string)context.InputParameters["fieldsecurityprofilename"];
            tracer.Trace("Security profile name: " + securityProfileName);
            string permissionfilter = string.Empty;
            Guid primaryEntityId = new Guid((string)context.InputParameters["entityid"]);
            tracer.Trace("Primary entity id: " + primaryEntityId.ToString());
            string entityName = context.InputParameters.Contains("entityname") ? (string)context.InputParameters["entityname"] : string.Empty;
            tracer.Trace("Entity name: " + entityName);
            string fieldName = context.InputParameters.Contains("fieldname") ? (string)context.InputParameters["fieldname"] : string.Empty;
            tracer.Trace("Field name: " + fieldName);
            if (!string.IsNullOrEmpty(entityName) || !string.IsNullOrEmpty(fieldName))
            {
                string conditions = string.IsNullOrEmpty(entityName) ? "" : @"<condition attribute='entityname' operator='eq' value='" + entityName + "'/>";
                conditions += string.IsNullOrEmpty(fieldName) ? "" : @"<condition attribute='attributelogicalname' operator='eq' value='" + fieldName + "'/>";
                permissionfilter = @"<filter type='and'>" + conditions + "</filter>";
            }
            tracer.Trace("Generating system user profiles query.");
            string query = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                              <entity name='systemuserprofiles'>
                                <attribute name='systemuserid' />
                                <filter type='and'>
                                  <condition attribute='systemuserid' operator='eq' value='" + primaryEntityId.ToString() + @"'/>
                                </filter>
                                <link-entity name='fieldsecurityprofile' from='fieldsecurityprofileid' to='fieldsecurityprofileid' link-type='inner' alias='fsp'>
                                  <filter type='and'>
                                    <condition attribute='name' operator='eq' value='" + securityProfileName + @"'/>
                                  </filter>
                                  <link-entity name='fieldpermission' from='fieldsecurityprofileid' to='fieldsecurityprofileid' alias='fp' link-type='inner'>
                                    <attribute name='canread' />
                                    <attribute name='attributelogicalname' />
                                    <attribute name='cancreate' />
                                    <attribute name='canupdate' />
                                    $fieldpermissionfilter$
                                  </link-entity>
                                </link-entity>
                              </entity>
                            </fetch>".Replace("$fieldpermissionfilter$", permissionfilter);

            EntityCollection result = service.RetrieveMultiple(new FetchExpression(query));
            tracer.Trace("Result count: " + result.Entities.Count);
            if (result.Entities.Count > 0)
            {
                foreach (Entity userProfile in result.Entities)
                {
                    if (userProfile.Contains("fp.attributelogicalname"))
                    {
                        string logicalName = (string)userProfile.GetAttributeValue<AliasedValue>("fp.attributelogicalname").Value;
                        FieldPermissions oFP = new FieldPermissions();
                        if (userProfile.Contains("fp.canread"))
                        {
                            oFP.canRead = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.canread").Value).Value;
                        }
                        if (userProfile.Contains("fp.cancreate"))
                        {
                            oFP.canCreate = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.cancreate").Value).Value;
                        }
                        if (userProfile.Contains("fp.canupdate"))
                        {
                            oFP.canUpdate = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.canupdate").Value).Value;
                        }
                        if (!sup.dctFieldPermissions.ContainsKey(logicalName))
                        {
                            sup.dctFieldPermissions.Add(logicalName, oFP);
                        }
                        else
                        {
                            if (oFP.canCreate == 4)
                            {
                                sup.dctFieldPermissions[logicalName].canCreate = 4;
                            }
                            if (oFP.canUpdate == 4)
                            {
                                sup.dctFieldPermissions[logicalName].canUpdate = 4;
                            }
                            if (oFP.canRead == 4)
                            {
                                sup.dctFieldPermissions[logicalName].canRead = 4;
                            }
                        }
                    }
                }
            }
            query = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                  <entity name='teamprofiles'>
                                    <attribute name='teamid' />
                                    <link-entity name='teammembership' from='teamid' to='teamid' link-type='inner' alias='tm'>
                                      <filter type='and'>
                                        <condition attribute='systemuserid' operator='eq' value='" + primaryEntityId.ToString() + @"'/>
                                      </filter>
                                    </link-entity>
                                    <link-entity name='fieldsecurityprofile' from='fieldsecurityprofileid' to='fieldsecurityprofileid' link-type='inner' alias='fsp'>
                                      <filter type='and'>
                                        <condition attribute='name' operator='eq' value='" + securityProfileName + @"'/> 
                                      </filter>
                                      <link-entity name='fieldpermission' from='fieldsecurityprofileid' to='fieldsecurityprofileid' alias='fp' link-type='inner'>
                                        <attribute name='attributelogicalname' />
                                        <attribute name='canread' />
                                        <attribute name='cancreate' />
                                        <attribute name='canupdate' />
                                        $fieldpermissionfilter$
                                      </link-entity>
                                    </link-entity>
                                  </entity>
                                </fetch>".Replace("$fieldpermissionfilter$", permissionfilter);
            result = service.RetrieveMultiple(new FetchExpression(query));
            tracer.Trace("Result count: " + result.Entities.Count);
            if (result.Entities.Count > 0)
            {
                foreach (Entity userProfile in result.Entities)
                {
                    if (userProfile.Contains("fp.attributelogicalname"))
                    {
                        string logicalName = (string)userProfile.GetAttributeValue<AliasedValue>("fp.attributelogicalname").Value;
                        FieldPermissions oFP = new FieldPermissions();
                        if (userProfile.Contains("fp.canread"))
                        {
                            oFP.canRead = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.canread").Value).Value;
                        }
                        if (userProfile.Contains("fp.cancreate"))
                        {
                            oFP.canCreate = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.cancreate").Value).Value;
                        }
                        if (userProfile.Contains("fp.canupdate"))
                        {
                            oFP.canUpdate = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.canupdate").Value).Value;
                        }
                        sup.dctFieldPermissions.Add(logicalName, oFP);
                        if (!sup.dctFieldPermissions.ContainsKey(logicalName))
                        {
                            sup.dctFieldPermissions.Add(logicalName, oFP);
                        }
                        else
                        {
                            if (oFP.canCreate == 4)
                            {
                                sup.dctFieldPermissions[logicalName].canCreate = 4;
                            }
                            if (oFP.canUpdate == 4)
                            {
                                sup.dctFieldPermissions[logicalName].canUpdate = 4;
                            }
                            if (oFP.canRead == 4)
                            {
                                sup.dctFieldPermissions[logicalName].canRead = 4;
                            }
                        }
                    }
                }
            }
            return sup;
        }

        public TeamProfile CheckTeamProfileAssociation(IPluginExecutionContext context, IOrganizationService service, ITracingService tracer)
        {
            TeamProfile tp = new TeamProfile();
            string securityProfileName = (string)context.InputParameters["fieldsecurityprofilename"];
            tracer.Trace("Security profile name: " + securityProfileName);
            string permissionfilter = string.Empty;
            Guid primaryEntityId = new Guid((string)context.InputParameters["entityid"]);
            tracer.Trace("Primary entity id: " + primaryEntityId.ToString());
            string entityName = context.InputParameters.Contains("entityname") ? (string)context.InputParameters["entityname"] : string.Empty;
            tracer.Trace("Entity name: " + entityName);
            string fieldName = context.InputParameters.Contains("fieldname") ? (string)context.InputParameters["fieldname"] : string.Empty;
            tracer.Trace("Field name: " + fieldName);
            if (!string.IsNullOrEmpty(entityName) || !string.IsNullOrEmpty(fieldName))
            {
                string conditions = string.IsNullOrEmpty(entityName) ? "" : @"<condition attribute='entityname' operator='eq' value='" + entityName + "'/>";
                conditions += string.IsNullOrEmpty(fieldName) ? "" : @"<condition attribute='attributelogicalname' operator='eq' value='" + fieldName + "'/>";
                permissionfilter = @"<filter type='and'>" + conditions + "</filter>";
            }
            tracer.Trace("Generating system user profiles query.");
            string query = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                              <entity name='teamprofiles'>
                                <attribute name='teamid' />
                                <filter type='and'>
                                  <condition attribute='teamid' operator='eq' value='" + primaryEntityId.ToString() + @"'/>
                                </filter>
                                <link-entity name='fieldsecurityprofile' from='fieldsecurityprofileid' to='fieldsecurityprofileid' link-type='inner' alias='fsp'>
                                  <filter type='and'>
                                    <condition attribute='name' operator='eq' value='" + securityProfileName + @"'/>
                                  </filter>
                                  <link-entity name='fieldpermission' from='fieldsecurityprofileid' to='fieldsecurityprofileid' alias='fp' link-type='inner'>
                                    <attribute name='attributelogicalname' />
                                    <attribute name='canread' />
                                    <attribute name='cancreate' />
                                    <attribute name='canupdate' />
                                    $fieldpermissionfilter$
                                  </link-entity>
                                </link-entity>
                              </entity>
                            </fetch>".Replace("$fieldpermissionfilter$", permissionfilter);

            EntityCollection result = service.RetrieveMultiple(new FetchExpression(query));
            tracer.Trace("Result count: " + result.Entities.Count);
            if (result.Entities.Count > 0)
            {
                foreach (Entity userProfile in result.Entities)
                {
                    if (userProfile.Contains("fp.attributelogicalname"))
                    {
                        string logicalName = (string)userProfile.GetAttributeValue<AliasedValue>("fp.attributelogicalname").Value;
                        if (!tp.dctFieldPermissions.ContainsKey(logicalName))
                        {
                            FieldPermissions oFP = new FieldPermissions();
                            if (userProfile.Contains("fp.canread"))
                            {
                                oFP.canRead = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.canread").Value).Value;
                            }
                            if (userProfile.Contains("fp.cancreate"))
                            {
                                oFP.canCreate = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.cancreate").Value).Value;
                            }
                            if (userProfile.Contains("fp.canupdate"))
                            {
                                oFP.canUpdate = ((OptionSetValue)userProfile.GetAttributeValue<AliasedValue>("fp.canupdate").Value).Value;
                            }
                            tp.dctFieldPermissions.Add(logicalName, oFP);
                        }
                    }
                }
            }
            return tp;
        }
    }

    public class SystemUserProfile
    {
        internal Dictionary<string, FieldPermissions> dctFieldPermissions;

        public SystemUserProfile()
        {
            dctFieldPermissions = new Dictionary<string, FieldPermissions>();
        }
    }

    public class TeamProfile
    {
        internal Dictionary<string, FieldPermissions> dctFieldPermissions;

        public TeamProfile()
        {
            dctFieldPermissions = new Dictionary<string, FieldPermissions>();
        }
    }

    public class FieldPermissions
    {
        internal int canCreate, canUpdate, canRead;

        public FieldPermissions()
        {
            canCreate = 0;
            canRead = 0;
            canUpdate = 0;
        }
    }
}
