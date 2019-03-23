using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Functions;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PublishingFunctionProcessor: BaseFunctionProcessor
    {
        private PublishingPageTransformation publishingPageTransformation;
        private List<AddOnType> addOnTypes;
        private object builtInFunctions;
        private ClientContext sourceClientContext;
        private ListItem page;

        #region Construction
        public PublishingFunctionProcessor(ListItem page, ClientContext sourceClientContext, PublishingPageTransformation publishingPageTransformation)
        {
            this.page = page;
            this.publishingPageTransformation = publishingPageTransformation;
            this.sourceClientContext = sourceClientContext;

            RegisterAddons();
        }
        #endregion

        #region Public methods
        public Tuple<string, string> Process(WebPartProperty webPartProperty)
        {
            string propertyKey = "";
            string propertyValue = "";

            if (!string.IsNullOrEmpty(webPartProperty.Functions))
            {
                var functionDefinition = ParseFunctionDefinition(webPartProperty.Functions, webPartProperty, this.page);

                // Execute function
                MethodInfo methodInfo = null;
                object functionClassInstance = null;

                if (string.IsNullOrEmpty(functionDefinition.AddOn))
                {
                    // Native builtin function
                    methodInfo = typeof(PublishingBuiltIn).GetMethod(functionDefinition.Name);
                    functionClassInstance = this.builtInFunctions;
                }
                else
                {
                    // Function specified via addon
                    var addOn = this.addOnTypes.Where(p => p.Name.Equals(functionDefinition.AddOn, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (addOn != null)
                    {
                        methodInfo = addOn.Type.GetMethod(functionDefinition.Name);
                        functionClassInstance = addOn.Instance;
                    }
                }

                if (methodInfo != null)
                {
                    // Execute the function
                    object result = ExecuteMethod(functionClassInstance, functionDefinition, methodInfo);

                    if (result is string || result is bool)
                    {
                        propertyKey = webPartProperty.Name;
                        propertyValue = result.ToString().ToLower();
                    }
                }
            }

            return new Tuple<string, string>(propertyKey, propertyValue);
        }
        #endregion

        #region Helper methods
        private static FunctionDefinition ParseFunctionDefinition(string function, WebPartProperty webPartProperty, ListItem page)
        {
            // Supported function syntax: 
            // - EncodeGuid()
            // - MyLib.EncodeGuid()
            // - EncodeGuid({ListId})
            // - EncodeGuid({ListId}, {Param2})
            // - {ViewId} = EncodeGuid()
            // - {ViewId} = EncodeGuid({ListId})
            // - {ViewId} = MyLib.EncodeGuid({ListId})
            // - {ViewId} = EncodeGuid({ListId}, {Param2})

            FunctionDefinition def = new FunctionDefinition();

            string functionString = null;
            if (function.IndexOf("=") > 0)
            {
                var split = function.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                FunctionParameter output = new FunctionParameter()
                {
                    Name = split[0].Replace("{", "").Replace("}", "").Trim(),
                    Type = FunctionType.String
                };

                def.Output = output;
                functionString = split[1].Trim();
            }
            else
            {
                FunctionParameter output = new FunctionParameter()
                {
                    Name = webPartProperty.Name,
                    Type = FunctionType.String
                };

                def.Output = output;
                functionString = function.Trim();
            }


            string functionName = functionString.Substring(0, functionString.IndexOf("("));
            if (functionName.IndexOf(".") > -1)
            {
                // This is a custom function
                def.AddOn = functionName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries)[0];
                def.Name = functionName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries)[1];
            }
            else
            {
                // this is an BuiltIn function
                def.AddOn = "";
                def.Name = functionString.Substring(0, functionString.IndexOf("("));
            }

            def.Input = new List<FunctionParameter>();

            var functionParameters = functionString.Substring(functionString.IndexOf("(") + 1).Replace(")", "").Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var functionParameter in functionParameters)
            {
                FunctionParameter input = new FunctionParameter()
                {
                    Name = functionParameter.Replace("{", "").Replace("}", "").Trim(),
                };

                // Populate the function parameter with a value coming from the analyzed web part
                input.Type = MapType("string");
                input.Value = page[input.Name].ToString();
                def.Input.Add(input);

                /*
                var wpProp = webPartData.Properties.Where(p => p.Name.Equals(input.Name, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefault();
                if (wpProp != null)
                {
                    // Map types used in the model to types used in function processor
                    input.Type = MapType(wpProp.Type.ToString());
                    var wpInstanceProp = webPart.Properties.Where(p => p.Key.Equals(input.Name, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefault();
                    input.Value = wpInstanceProp.Value;
                    def.Input.Add(input);
                }
                else
                {
                    throw new Exception($"Parameter {input.Name} was used but is not listed as a web part property that can be used.");
                }
                */
            }

            return def;
        }


        private void RegisterAddons()
        {
            // instantiate default built in functions class
            this.addOnTypes = new List<AddOnType>();
            this.builtInFunctions = Activator.CreateInstance(typeof(PublishingBuiltIn), sourceClientContext, base.RegisteredLogObservers);

            // instantiate the custom function classes (if there are)
            if (this.publishingPageTransformation.AddOns != null)
            {
                foreach (var addOn in this.publishingPageTransformation.AddOns)
                {
                    try
                    {
                        string path = "";
                        if (addOn.Assembly.Contains("\\") && System.IO.File.Exists(addOn.Assembly))
                        {
                            path = addOn.Assembly;
                        }
                        else
                        {
                            path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, addOn.Assembly);
                        }

                        var assembly = Assembly.LoadFile(path);
                        var customType = assembly.GetType(addOn.Type);
                        var instance = Activator.CreateInstance(customType, sourceClientContext);

                        this.addOnTypes.Add(new AddOnType()
                        {
                            Name = addOn.Name,
                            Assembly = assembly,
                            Instance = instance,
                            Type = customType,
                        });
                    }
                    catch (Exception ex)
                    {
                        LogError(LogStrings.Error_FailedToInitiateCustomFunctionClasses, LogStrings.Heading_FunctionProcessor, ex);
                        throw;
                    }
                }
            }
        }
        #endregion

    }
}
