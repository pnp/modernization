using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Functions;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PublishingFunctionProcessor: BaseFunctionProcessor
    {
        public enum FieldType
        {
            String = 0,
            Bool = 1,
            Guid = 2,
            Integer = 3,
            DateTime = 4,
            User = 5,
        }

        private PublishingPageTransformation publishingPageTransformation;
        private List<AddOnType> addOnTypes;
        private object builtInFunctions;
        private ClientContext sourceClientContext;
        private ClientContext targetClientContext;
        private ListItem page;

        #region Construction
        public PublishingFunctionProcessor(ListItem page, ClientContext sourceClientContext, ClientContext targetClientContext, PublishingPageTransformation publishingPageTransformation)
        {
            this.page = page;
            this.publishingPageTransformation = publishingPageTransformation;
            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;

            RegisterAddons();
        }
        #endregion

        #region Public methods
        //public Tuple<string, string> Process(WebPartProperty webPartProperty)
        public Tuple<string, string> Process(string functions, string propertyName, FieldType propertyType)
        {
            string propertyKey = "";
            string propertyValue = "";

            if (!string.IsNullOrEmpty(functions))
            {
                var functionDefinition = ParseFunctionDefinition(functions, propertyName, propertyType, this.page);

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

                    // output types support: string or bool
                    if (result is string || result is bool)
                    {
                        //propertyKey = webPartProperty.Name;
                        propertyKey = propertyName;
                        propertyValue = result.ToString().ToLower();
                    }
                }
            }

            return new Tuple<string, string>(propertyKey, propertyValue);
        }
        #endregion

        #region Helper methods
        //private static FunctionDefinition ParseFunctionDefinition(string function, WebPartProperty webPartProperty, ListItem page)
        private static FunctionDefinition ParseFunctionDefinition(string function, string propertyName, FieldType propertyType, ListItem page)
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
                    Name = propertyName,
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

                // Populate the function parameter with a value coming from publishing page
                input.Type = MapType(propertyType.ToString());

                if (propertyType == FieldType.String)
                {
                    input.Value = page.GetFieldValueAs<string>(input.Name);
                }
                else if (propertyType == FieldType.User)
                {
                    if (page.FieldExistsAndUsed(input.Name))
                    {
                        input.Value = ((FieldUserValue)page[input.Name]).LookupId.ToString();
                    }
                }
                def.Input.Add(input);
            }

            return def;
        }



        private void RegisterAddons()
        {
            // instantiate default built in functions class
            this.addOnTypes = new List<AddOnType>();
            this.builtInFunctions = Activator.CreateInstance(typeof(PublishingBuiltIn), sourceClientContext, targetClientContext, base.RegisteredLogObservers);

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
