using Microsoft.Xrm.Sdk;
using R.iT.UtilityClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace biohof
{
    public class ArbeitsschrittAufInaktivStellen : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService proxy = serviceFactory.CreateOrganizationService(context.UserId);
            //ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if(context.Depth > 1) return;

            Entity targetEntity = (Entity)context.InputParameters["Target"];
            Entity preImage = context.MessageName == "Update" ? context.PreEntityImages["record"] : targetEntity;
            Entity postEntity = CommonMethods.AssemblePostImage(targetEntity, preImage);

            DateTime processDate = postEntity.GetAttributeValue<DateTime>("rit_processdate");

            if (processDate == null) return;

            if (processDate.Date < DateTime.Today.Date)
            {
                postEntity.Attributes["statecode"] = new OptionSetValue(1);
                postEntity.Attributes["statuscode"] = new OptionSetValue(2);
            }

            if (processDate.Date > DateTime.Today.Date)
            {
                postEntity.Attributes["statecode"] = new OptionSetValue(0);
                postEntity.Attributes["statuscode"] = new OptionSetValue(1);
            }




            proxy.Update(postEntity);
           
        }
    }
}
