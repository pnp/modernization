using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Transform
{
    public  class TermTransformator : BaseTransform
    {
        private ClientContext _sourceContext;
        private ClientContext _targetContext;
        private List<TermMapping> termMapping;
        private bool skipDefaultTermStoreMapping;
        private string TermNodeDelimiter = "|";

        #region Construction        

        /// <summary>
        /// Constructor for the Term Transformator class
        /// </summary>
        /// <param name="baseTransformationInformation"></param>
        /// <param name="sourceContext"></param>
        /// <param name="targetContext"></param>
        /// <param name="logObservers"></param>
        public TermTransformator(BaseTransformationInformation baseTransformationInformation, ClientContext sourceContext, ClientContext targetContext, IList<ILogObserver> logObservers = null)
        {
            // Hookup logging
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            // Ensure source and target context are set
            if (sourceContext == null && targetContext != null)
            {
                sourceContext = targetContext;
            }

            if (targetContext == null && sourceContext != null)
            {
                targetContext = sourceContext;
            }


            this._sourceContext = sourceContext;
            this._targetContext = targetContext;

            // Load the Term mapping file
            if (!string.IsNullOrEmpty(baseTransformationInformation?.TermMappingFile))
            {
                this.termMapping = CacheManager.Instance.GetTermMapping(baseTransformationInformation.TermMappingFile, logObservers);
            }

            this.skipDefaultTermStoreMapping = baseTransformationInformation.SkipTermStoreMapping;
        }

        #endregion

        /// <summary>
        /// Transforms a collection of terms in a dictionary
        /// </summary>
        /// <returns></returns>
        public TaxonomyFieldValueCollection TransformCollection(TaxonomyFieldValueCollection taxonomyFieldValueCollection)
        {
            throw new NotImplementedException();
            
        }

        /// <summary>
        /// Transforms a collection of terms in a dictionary
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, object> TransformCollection(Dictionary<string, object> fieldValueCollection)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Main entry method for transforming terms
        /// </summary>
        public TermData Transform(TermData inputSourceTerm)
        {
            //Design:
            // This will have two modes:
            // Default mode to work out the terms from source to destination based on identical IDs or Term Paths
            // Mapping file to override default mode for specifically mapping a source term to designation term

            //Scenarios:
            // Term Ids or Term Names
            // Source or Target Term ID/Name may not be found



            if (this.skipDefaultTermStoreMapping)
            {
                // Mapping mode only
                if (termMapping != null)
                {

                }
            }
            else
            {
                // Default Mode 


                // Mapping Mode 
                if (termMapping != null)
                {

                }
            }
                       
            return inputSourceTerm; //Pass-Through
        }

        
        

        public void ValidateSourceTerm() {
            throw new NotImplementedException();
        }

        public void ValidateTargetTerm()
        {
            throw new NotImplementedException();
        }

        public bool IsValidGuid()
        {
            throw new NotImplementedException();
        }
    }
}
