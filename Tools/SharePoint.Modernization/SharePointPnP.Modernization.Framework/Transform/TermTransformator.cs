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
    public class TermTransformator : BaseTransform
    {
        private ClientContext _sourceContext;
        private ClientContext _targetContext;
        private List<TermMapping> termMapping;
        private bool skipDefaultTermStoreMapping;

        public const string TermNodeDelimiter = "|";

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

            if (baseTransformationInformation != null)
            {
                this.skipDefaultTermStoreMapping = baseTransformationInformation.SkipTermStoreMapping;
            }
        }

        #endregion

        /// <summary>
        /// Transforms a collection of terms in a dictionary
        /// </summary>
        /// <returns></returns>
        public TaxonomyFieldValueCollection TransformCollection(TaxonomyFieldValueCollection taxonomyFieldValueCollection)
        {
            foreach (var fieldValue in taxonomyFieldValueCollection)
            {
                var result = this.Transform(new TermData() { TermGuid = Guid.Parse(fieldValue.TermGuid), TermLabel = fieldValue.Label });
                fieldValue.Label = result.TermLabel;
                fieldValue.TermGuid = result.TermLabel;
            }

            return taxonomyFieldValueCollection;
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

        /// <summary>
        /// Sets the cache for contents of the term store to be used when getting terms for fields
        /// </summary>
        /// <param name="termSetId"></param>
        /// <param name="isSourceTermStore"></param>
        public void CacheTermsFromTermStore(Guid sourceTermSetId, Guid targetTermSetId)
        {
            // Collect source terms
            if (sourceTermSetId != null && sourceTermSetId != Guid.Empty)
            {
                Cache.CacheManager.Instance.StoreTermSetTerms(this._sourceContext, sourceTermSetId);
            }

            if (targetTermSetId != null && targetTermSetId != Guid.Empty)
            {
                Cache.CacheManager.Instance.StoreTermSetTerms(this._sourceContext, targetTermSetId);
            }

        }
        
        /// <summary>
        /// Extract all the terms from a termset for caching and quicker processing
        /// </summary>
        /// <param name="termSetId"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public static Dictionary<Guid, TermData> GetAllTermsFromTermSet(Guid termSetId, ClientContext context)
        {
            //Use a source and target Dictionary<guid,string>
            //Key = Id, Value = Path(e.g.termgroup | termset | term)
            var termsCache = new Dictionary<Guid, TermData>();

            try
            {
                TaxonomySession session = TaxonomySession.GetTaxonomySession(context);
                TermStore termStore = session.GetDefaultSiteCollectionTermStore();
                var termSet = termStore.GetTermSet(termSetId);
                var termGroup = termSet.Group;
                context.Load(termSet, t => t.Terms, t => t.Name);
                context.Load(termGroup, g => g.Name);
                context.ExecuteQueryRetry();

                var termGroupName = termGroup.Name;
                var setName = termSet.Name;
                var termSetPath = $"{termGroupName}{TermTransformator.TermNodeDelimiter}{setName}";
                foreach (var term in termSet.Terms)
                {
                    var termName = term.Name;
                    var termPath = $"{termSetPath}{TermNodeDelimiter}{termName}";
                    termsCache.Add(term.Id, 
                        new TermData() { TermGuid = term.Id, TermLabel = termName, TermPath = termPath, TermSetId = termSetId });

                    if (term.TermsCount > 0)
                    {
                        var subTerms = ParseSubTerms(termPath, term, termSetId, context);
                        //termsCache
                        foreach (var foundTerm in subTerms)
                        {
                            termsCache.Add(foundTerm.Key, foundTerm.Value);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                //TODO: Record any failure
            }

            return termsCache;
        }

        /// <summary>
        /// Gets the term labels within a term recursively
        /// </summary>
        /// <param name="subTermPath"></param>
        /// <param name="term"></param>
        /// <param name="includeId"></param>
        /// <param name="delimiter"></param>
        /// <param name="clientContext"></param>
        /// <returns></returns>
        /// Reference: https://github.com/SharePoint/PnP-Sites-Core/blob/master/Core/OfficeDevPnP.Core/Extensions/TaxonomyExtensions.cs
        public static Dictionary<Guid, TermData> ParseSubTerms(string subTermPath, Term term, Guid termSetId, ClientRuntimeContext clientContext)
        {
            var items = new Dictionary<Guid, TermData>();
            if (term.ServerObjectIsNull == null || term.ServerObjectIsNull == false)
            {
                clientContext.Load(term.Terms);
                clientContext.ExecuteQueryRetry();
            }

            foreach (var subTerm in term.Terms)
            {
                var termName = subTerm.Name;
                var termPath = $"{subTermPath}{TermTransformator.TermNodeDelimiter}{termName}";

                items.Add(subTerm.Id, new TermData() { TermGuid = subTerm.Id, TermLabel = termName, TermPath = termPath, TermSetId = termSetId });

                if (term.TermsCount > 0)
                {
                    var moreSubTerms = ParseSubTerms(termPath, subTerm, termSetId, clientContext);
                    foreach(var foundTerm in moreSubTerms)
                    {
                        items.Add(foundTerm.Key, foundTerm.Value);
                    }
                }

            }
            return items;
        }

        public void ValidateSourceTerm()
        {
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
