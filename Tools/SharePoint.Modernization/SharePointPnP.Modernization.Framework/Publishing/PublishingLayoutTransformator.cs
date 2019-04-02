using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    /// <summary>
    /// Specific layout transformator for the 'AutoDetect' layout option for publishing pages
    /// </summary>
    public class PublishingLayoutTransformator : BaseTransform, ILayoutTransformator
    {
        private ClientSidePage page;

        #region Construction
        /// <summary>
        /// Creates a layout transformator instance
        /// </summary>
        /// <param name="page">Client side page that will be receive the created layout</param>
        public PublishingLayoutTransformator(ClientSidePage page, IList<ILogObserver> logObservers = null)
        {
            // Register observers
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            this.page = page;
        }
        #endregion


        public void Transform(Tuple<Pages.PageLayout, List<WebPartEntity>> pageData)
        {
            // First drop all sections...ensure the default section is gone
            page.Sections.Clear();

            // Should not occur, but to be at the safe side...
            if (pageData.Item2.Count == 0)
            {
                page.AddSection(CanvasSectionTemplate.OneColumn, 1);
                return;
            }

            var firstRow = pageData.Item2.OrderBy(p => p.Row).First().Row;
            var lastRow = pageData.Item2.OrderBy(p => p.Row).Last().Row;

            // Loop over the possible rows...will take in account possible row gaps
            // Each row means a new section
            int sectionOrder = 1;
            for (int rowIterator = firstRow; rowIterator <= lastRow; rowIterator++)
            {
                var webpartsInRow = pageData.Item2.Where(p => p.Row == rowIterator);
                if (webpartsInRow.Any())
                {
                    // Determine max column number
                    int maxColumns = 1;

                    foreach (var wpInRow in webpartsInRow)
                    {
                        if (wpInRow.Column > maxColumns)
                        {
                            maxColumns = wpInRow.Column;
                        }
                    }

                    if (maxColumns > 3)
                    {
                        LogError(LogStrings.Error_Maximum3ColumnsAllowed, LogStrings.Heading_PublishingLayoutTransformator);
                        throw new Exception("Publishing transformation layout mapping can maximum use 3 columns");
                    }
                    else
                    {
                        if (maxColumns == 1)
                        {
                            page.AddSection(CanvasSectionTemplate.OneColumn, sectionOrder);
                        }
                        else if (maxColumns == 2)
                        {
                            // if we've only an image in one of the columns then make that one the 'small' column
                            var imageWebPartsInRow = webpartsInRow.Where(p => p.Type == WebParts.WikiImage);
                            if (imageWebPartsInRow.Any())
                            {
                                Dictionary<int, int> imageWebPartsPerColumn = new Dictionary<int, int>();
                                foreach(var imageWebPart in imageWebPartsInRow.OrderBy(p=>p.Column))
                                {
                                    if (imageWebPartsPerColumn.TryGetValue(imageWebPart.Column, out int wpCount))
                                    {
                                        imageWebPartsPerColumn[imageWebPart.Column] = wpCount + 1;
                                    }
                                    else
                                    {
                                        imageWebPartsPerColumn.Add(imageWebPart.Column, 1);
                                    }
                                }

                                var firstImageColumn = imageWebPartsPerColumn.First();
                                var secondImageColumn = imageWebPartsPerColumn.Last();

                                if (firstImageColumn.Key == secondImageColumn.Key)
                                {
                                    // there was only one column with images
                                    var firstImageColumnOtherWebParts = webpartsInRow.Where(p => p.Column == firstImageColumn.Key && p.Type != WebParts.WikiImage);
                                    if (firstImageColumnOtherWebParts.Count()==0)
                                    {
                                        // no other web parts in this column
                                        var orderedList = webpartsInRow.OrderBy(p => p.Column).First();

                                        if (orderedList.Column == firstImageColumn.Key)
                                        {
                                            // image left
                                            page.AddSection(CanvasSectionTemplate.TwoColumnRight, sectionOrder);
                                        }
                                        else
                                        {
                                            // image right
                                            page.AddSection(CanvasSectionTemplate.TwoColumnLeft, sectionOrder);
                                        }
                                    }
                                    else
                                    {
                                        page.AddSection(CanvasSectionTemplate.TwoColumn, sectionOrder);
                                    }
                                }
                                else
                                {
                                    if (firstImageColumn.Value == 1 || secondImageColumn.Value == 1)
                                    {
                                        // does one of the two columns have anything else besides image web parts
                                        var firstImageColumnOtherWebParts = webpartsInRow.Where(p => p.Column == firstImageColumn.Key && p.Type != WebParts.WikiImage);
                                        var secondImageColumnOtherWebParts = webpartsInRow.Where(p => p.Column == secondImageColumn.Key && p.Type != WebParts.WikiImage);

                                        if (firstImageColumnOtherWebParts.Count() == 0 && secondImageColumnOtherWebParts.Count() == 0)
                                        {
                                            // two columns with each only one image...
                                            page.AddSection(CanvasSectionTemplate.TwoColumn, sectionOrder);
                                        }
                                        else if (firstImageColumnOtherWebParts.Count() == 0 && secondImageColumnOtherWebParts.Count() > 0)
                                        {
                                            page.AddSection(CanvasSectionTemplate.TwoColumnRight, sectionOrder);
                                        }
                                        else if (firstImageColumnOtherWebParts.Count() > 0 && secondImageColumnOtherWebParts.Count() == 0)
                                        {
                                            page.AddSection(CanvasSectionTemplate.TwoColumnLeft, sectionOrder);
                                        }
                                        else
                                        {
                                            page.AddSection(CanvasSectionTemplate.TwoColumn, sectionOrder);
                                        }
                                    }
                                    else
                                    {
                                        page.AddSection(CanvasSectionTemplate.TwoColumn, sectionOrder);
                                    }
                                }
                            }
                            else
                            {
                                page.AddSection(CanvasSectionTemplate.TwoColumn, sectionOrder);
                            }
                        }
                        else if (maxColumns == 3)
                        {
                            page.AddSection(CanvasSectionTemplate.ThreeColumn, sectionOrder);
                        }

                        sectionOrder++;
                    }
                }
                else
                {
                    // non used row...ignore
                }
            }
        }
    }
}
