using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using Oas2xlsx.Console.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Oas2xlsx.Console
{
    public class ExcelGenerator
    {
        private readonly OpenApiDocument _oasDocument;

        public ExcelGenerator(OpenApiDocument oasDocument)
        {
            if (oasDocument == null)
            {
                throw new ArgumentNullException("oasDocument");
            }
            _oasDocument = oasDocument;
        }

        public void Generate(string outputPath)
        {
            using (var excelDocument = new XLWorkbook())
            {
                GeneratePaths(excelDocument);


                using (var stream = new MemoryStream())
                {
                    excelDocument.SaveAs(stream);
                    var content = stream.ToArray();

                    try
                    {
                        File.WriteAllBytes(outputPath, content);
                    }
                    catch (System.IO.IOException e)
                    {
                        ColorConsole.WriteError(e.Message);
                        return;
                    }
                }
            }
        }

        private void GeneratePaths(XLWorkbook excelDocument)
        {
            foreach (var path in _oasDocument.Paths)
            {
                // Be sure that the worksheet title will have a length of 31 character maximum
                int worksheetNameStart = ((int)path.Key.Length - 31) >= 0 ? (int)path.Key.Length - 24 : 0;
                string worksheetName = (worksheetNameStart == 0) ?
                                            path.Key :
                                            string.Format("(...){0}", path.Key.Substring(worksheetNameStart, 24));
                // Replace worksheet unsupported characters
                worksheetName = worksheetName.Replace("/", "#").Replace(":", "#");
                // if the worksheet name is already presents, we add an increment
                int worksheetNumber = 1;
                while (excelDocument.Worksheets.Contains(worksheetName))
                {
                    worksheetName = string.Format("{0}{1}", worksheetName, worksheetNumber);
                    worksheetNumber++;
                }
                var worksheet = excelDocument.Worksheets.Add(worksheetName);

                var currentRow = 1;

                worksheet.Row(currentRow).Row(1, 10).Style.Fill.BackgroundColor = XLColor.OrangeRed;
                worksheet.Cell(currentRow, 1).Value = "Full path name: ";
                worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                worksheet.Cell(currentRow, 1).Style.Font.FontColor = XLColor.White;
                worksheet.Cell(currentRow, 2).Value = path.Key;
                worksheet.Range(currentRow, 2, currentRow, 10).Merge();
                worksheet.Cell(currentRow, 2).Style.Font.Bold = true;
                worksheet.Cell(currentRow, 1).Style.Font.FontColor = XLColor.White;

                if (!string.IsNullOrEmpty(path.Value.Description))
                {
                    worksheet.Cell(currentRow, 2).Comment.AddNewLine().AddText(path.Value.Description);
                    worksheet.Cell(currentRow, 2).Comment.Style.Size.SetAutomaticSize();
                }
                currentRow++;

                currentRow++; // add a new line to easy readability
                foreach (var operation in path.Value.Operations)
                {
                    worksheet.Row(currentRow).Row(1, 10).Style.Fill.BackgroundColor = XLColor.DarkOrange;
                    worksheet.Cell(currentRow, 1).Value = "Supported operation: ";
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Font.FontColor = XLColor.White;
                    worksheet.Cell(currentRow, 2).Value = operation.Key;
                    worksheet.Cell(currentRow, 2).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Font.FontColor = XLColor.White;
                    if (!string.IsNullOrEmpty(operation.Value.Description))
                    {
                        worksheet.Cell(currentRow, 2).Comment.AddNewLine().AddText(operation.Value.Description);
                        worksheet.Cell(currentRow, 2).Comment.Style.Size.SetAutomaticSize();
                    }
                    currentRow++;

                    currentRow++; // add a new line to easy readability
                    switch (operation.Key)
                    {
                        case OperationType.Get:
                            GenerateHeaders(worksheet, operation.Value, ref currentRow);
                            GenerateQueryParams(worksheet, operation.Value, ref currentRow);
                            GenerateResponseBody(worksheet, operation.Value, ref currentRow);
                            break;
                        case OperationType.Put:
                            GenerateHeaders(worksheet, operation.Value, ref currentRow);
                            GenerateRequestBody(worksheet, operation.Value, ref currentRow);
                            GenerateResponseBody(worksheet, operation.Value, ref currentRow);
                            break;
                        case OperationType.Post:
                            GenerateHeaders(worksheet, operation.Value, ref currentRow);
                            GenerateRequestBody(worksheet, operation.Value, ref currentRow);
                            GenerateResponseBody(worksheet, operation.Value, ref currentRow);
                            break;
                        case OperationType.Delete:
                            GenerateHeaders(worksheet, operation.Value, ref currentRow);
                            GenerateResponseBody(worksheet, operation.Value, ref currentRow);
                            break;
                        case OperationType.Patch:
                            GenerateHeaders(worksheet, operation.Value, ref currentRow);
                            GenerateRequestBody(worksheet, operation.Value, ref currentRow);
                            GenerateResponseBody(worksheet, operation.Value, ref currentRow);
                            break;
                        case OperationType.Options:
                            throw new NotImplementedException("Trace operation is not supported");
                        case OperationType.Head:
                            throw new NotImplementedException("Trace operation is not supported");
                        case OperationType.Trace:
                            throw new NotImplementedException("Trace operation is not supported");
                        default:
                            throw new NotImplementedException("Unknown operation: " + operation.Key);
                    }

                    currentRow++;
                }

                worksheet.Columns(1, 6).AdjustToContents(30, 120);

            }
        }

        /// <summary>
        /// Generate info about the available query parameters of a given operation
        /// </summary>
        /// <param name="worksheet">The worksheet where data have to be written</param>
        /// <param name="currentOperation">The current HTTP operation described in the OAS definition</param>
        /// <param name="currentRow">Current row of the worksheet, where data have to be written</param>
        private void GenerateQueryParams(IXLWorksheet worksheet, OpenApiOperation currentOperation, ref int currentRow)
        {
            int firstTableQueryParamsCellRow = currentRow;
            worksheet.Cell(currentRow, 1).Value = "Query parameter name";
            worksheet.Cell(currentRow, 2).Value = "Query parameter type";
            worksheet.Cell(currentRow, 3).Value = "Is mandatory?";
            currentRow++;


            if (currentOperation.Parameters.Any(q => q.In == ParameterLocation.Query) == false)
            {
                worksheet.Range(currentRow, 1, currentRow, 3).Merge();
                worksheet.Cell(currentRow, 1).Value = "No query param";
                currentRow++;
            }
            else
            {
                foreach (var queryParam in currentOperation.Parameters.Where(q => q.In == ParameterLocation.Query))
                {
                    worksheet.Cell(currentRow, 1).Value = queryParam.Name;
                    worksheet.Cell(currentRow, 2).Value = queryParam.Schema.Type;
                    worksheet.Cell(currentRow, 3).Value = queryParam.Required;
                    if (!string.IsNullOrEmpty(queryParam.Description))
                    {
                        worksheet.Cell(currentRow, 1).Comment.AddNewLine().AddText(queryParam.Description);
                        worksheet.Cell(currentRow, 1).Comment.Style.Size.SetAutomaticSize();
                    }
                    currentRow++;
                }
            }

            int lastTableQueryParamsCellRow = currentRow - 1;  // -1 to position the cursor on the last row of the table
            // Apply table style
            var tableQueryParams = worksheet.Range(firstTableQueryParamsCellRow, 1, lastTableQueryParamsCellRow, 4);
            var tableQueryParamsHeaders = tableQueryParams.Range(1, 1, 1, tableQueryParams.ColumnCount()); // coordonates are relative to the table
            tableQueryParamsHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            tableQueryParamsHeaders.Style.Font.Bold = true;
            tableQueryParamsHeaders.Style.Fill.BackgroundColor = XLColor.Orange;
            tableQueryParams.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            tableQueryParams.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            currentRow++; // add a new line to easy readability
        }

        /// <summary>
        /// Generate info about the available query headers of a given operation
        /// </summary>
        /// <param name="worksheet">The worksheet where data have to be written</param>
        /// <param name="currentOperation">The current HTTP operation described in the OAS definition</param>
        /// <param name="currentRow">Current row of the worksheet, where data have to be written</param>
        private void GenerateHeaders(IXLWorksheet worksheet, OpenApiOperation currentOperation, ref int currentRow)
        {
            int firstTableHeadersCellRow = currentRow;
            worksheet.Cell(currentRow, 1).Value = "Header name";
            worksheet.Cell(currentRow, 2).Value = "Header type";
            worksheet.Cell(currentRow, 3).Value = "Is mandatory?";
            currentRow++;

            if (currentOperation.Parameters.Any(q => q.In == ParameterLocation.Header) == false)
            {
                worksheet.Range(currentRow, 1, currentRow, 3).Merge();
                worksheet.Cell(currentRow, 1).Value = "No header";
                currentRow++;
            }
            else
            {
                foreach (var header in currentOperation.Parameters.Where(q => q.In == ParameterLocation.Header))
                {
                    worksheet.Cell(currentRow, 1).Value = header.Name;
                    worksheet.Cell(currentRow, 2).Value = header.Schema.Type;
                    worksheet.Cell(currentRow, 3).Value = header.Required;
                    if (!string.IsNullOrEmpty(header.Description))
                    {
                        worksheet.Cell(currentRow, 1).Comment.AddNewLine().AddText(header.Description);
                        worksheet.Cell(currentRow, 1).Comment.Style.Size.SetAutomaticSize();
                    }
                    currentRow++;
                }
            }

            int lastTableHeadersCellRow = currentRow - 1;  // -1 to position the cursor on the last row of the table
            // Apply table style
            var tableHeaders = worksheet.Range(firstTableHeadersCellRow, 1, lastTableHeadersCellRow, 3);
            var tableHeadersHeaders = tableHeaders.Range(1, 1, 1, tableHeaders.ColumnCount()); // coordonates are relative to the table
            tableHeadersHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            tableHeadersHeaders.Style.Font.Bold = true;
            tableHeadersHeaders.Style.Fill.BackgroundColor = XLColor.Orange;
            tableHeadersHeaders.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            tableHeadersHeaders.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            currentRow++; // add a new line to easy readability
        }

        /// <summary>
        /// Generate info about the request of a given operation
        /// </summary>
        /// <param name="worksheet">The worksheet where data have to be written</param>
        /// <param name="currentOperation">The current HTTP operation described in the OAS definition</param>
        /// <param name="currentRow">Current row of the worksheet, where data have to be written</param>
        private void GenerateRequestBody(IXLWorksheet worksheet, OpenApiOperation currentOperation, ref int currentRow)
        {
            if (currentOperation.RequestBody == null
                || currentOperation.RequestBody.Content.Count == 0)
            {
                return;
            }
            var requestBody = currentOperation.RequestBody.Content.FirstOrDefault();

            int firstTableRequestCellRow = currentRow;
            int firstTableRequestCellColumn = 1;

            worksheet.Cell(currentRow, 1).Value = "Request body format";
            worksheet.Cell(currentRow, 2).Value = requestBody.Key;
            worksheet.Range(currentRow, 2, currentRow, 3).Merge();
            currentRow++;

            worksheet.Cell(currentRow, 1).Value = "Request body type";
            worksheet.Cell(currentRow, 2).Value = requestBody.Value.Schema.Type;
            worksheet.Range(currentRow, 2, currentRow, 3).Merge();
            currentRow++;

            int lastTableRequestCellRow = currentRow - 1; // -1 to position the cursor on the last row of the table
            int lastTableRequestCellColumn = 3;

            // Apply table style
            var tableRequest = worksheet.Range(firstTableRequestCellRow, firstTableRequestCellColumn, lastTableRequestCellRow, lastTableRequestCellColumn);
            var tableRequestHeaders = tableRequest.Range(1, 1, tableRequest.RowCount(), 1); // coordonates are relative to the table
            tableRequestHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            tableRequestHeaders.Style.Font.Bold = true;
            tableRequestHeaders.Style.Fill.BackgroundColor = XLColor.Orange;
            tableRequest.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            tableRequest.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;


            int firstTableFieldsCellRow = currentRow;
            int firstTableFieldsCellColumn = 1;
            worksheet.Cell(currentRow, 1).Value = "Field name";
            worksheet.Cell(currentRow, 2).Value = "Field type";
            worksheet.Cell(currentRow, 3).Value = "Is mandatory?";
            currentRow++;
            foreach (var property in requestBody.Value.Schema.Properties)
            {
                var isRequired = requestBody.Value.Schema.Required.Contains(property.Key);
                GenerateProperty(worksheet, requestBody.Value.Schema, property.Key, property.Value, ref currentRow, isRequired);
            }
            // display sub objects
            foreach (var schema in requestBody.Value.Schema.AllOf)
            {
                foreach (var property in schema.Properties)
                {
                    var isRequired = schema.Required.Contains(property.Key);
                    GenerateProperty(worksheet, requestBody.Value.Schema, property.Key, property.Value, ref currentRow, isRequired);
                }
            }
            int oneOfIndex = 1;
            foreach (var schema in requestBody.Value.Schema.OneOf)
            {
                worksheet.Cell(currentRow, 1).Value = string.Format("Request can be of this schema #{0}", oneOfIndex);
                worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                worksheet.Cell(currentRow, 1).Style.Font.Italic = true;
                worksheet.Range(currentRow, 1, currentRow, 3).Merge();
                currentRow++;
                foreach (var property in schema.Properties)
                {
                    var isPropertyRequired = schema.Required.Contains(property.Key);
                    GenerateProperty(worksheet, requestBody.Value.Schema, property.Key, property.Value, ref currentRow, isPropertyRequired);
                }
                oneOfIndex++;
            }
            // display array item type
            if (requestBody.Value.Schema.Items != null)
            {
                // this property node doesn't need to be displayed, so we will skip it
                GenerateProperty(worksheet, requestBody.Value.Schema, string.Empty, requestBody.Value.Schema.Items, ref currentRow, true, true);
            }

            int lastTableFieldsCellRow = currentRow - 1; // -1 to position the cursor on the last row of the table
            int lastTableFieldsCellColumn = 3;

            // Apply table style
            var tableFields = worksheet.Range(firstTableFieldsCellRow, firstTableFieldsCellColumn, lastTableFieldsCellRow, lastTableFieldsCellColumn);
            var tableFieldsHeaders = tableFields.Range(1, 1, 1, tableFields.ColumnCount()); // coordonates are relative to the table
            tableFieldsHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            tableFieldsHeaders.Style.Font.Bold = true;
            tableFieldsHeaders.Style.Fill.BackgroundColor = XLColor.Orange;
            tableFields.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            tableFields.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            currentRow++; // add a new line to easy readability
        }

        /// <summary>
        /// Generate info about the available responses of a given operation
        /// </summary>
        /// <param name="worksheet">The worksheet where data have to be written</param>
        /// <param name="currentOperation">The current HTTP operation described in the OAS definition</param>
        /// <param name="currentRow">Current row of the worksheet, where data have to be written</param>
        private void GenerateResponseBody(IXLWorksheet worksheet, OpenApiOperation currentOperation, ref int currentRow)
        {
            if (currentOperation.Responses == null
                || currentOperation.Responses.Count == 0)
            {
                return;
            }

            foreach (var response in currentOperation.Responses)
            {
                int firstTableResponseCellRow = currentRow;
                int firstTableResponseCellColumn = 1;
                worksheet.Cell(currentRow, 1).Value = "Response HTTP code";
                worksheet.Cell(currentRow, 2).Value = response.Key;
                worksheet.Range(currentRow, 2, currentRow, 3).Merge();
                if (!string.IsNullOrEmpty(response.Value.Description))
                {
                    worksheet.Cell(currentRow, 2).Comment.AddNewLine().AddText(response.Value.Description);
                    worksheet.Cell(currentRow, 2).Comment.Style.Size.SetAutomaticSize();
                }
                currentRow++;

                KeyValuePair<string, OpenApiMediaType>? responseBody = null;
                if (response.Value.Content.Count > 0)
                {
                    responseBody = response.Value.Content.FirstOrDefault();
                    worksheet.Cell(currentRow, 1).Value = "Response body format";
                    worksheet.Cell(currentRow, 2).Value = responseBody.Value.Key;
                    worksheet.Range(currentRow, 2, currentRow, 3).Merge();
                    currentRow++;
                }

                int lastTableResponseCellRow = currentRow - 1; // -1 to position the cursor on the last row of the table
                int lastTableResponseCellColumn = 3;

                // Apply table style
                var tableResponse = worksheet.Range(firstTableResponseCellRow, firstTableResponseCellColumn, lastTableResponseCellRow, lastTableResponseCellColumn);
                var tableResponseHeaders = tableResponse.Range(1, 1, tableResponse.RowCount(), 1); // coordonates are relative to the table
                tableResponseHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                tableResponseHeaders.Style.Font.Bold = true;
                tableResponseHeaders.Style.Fill.BackgroundColor = XLColor.Orange;
                tableResponse.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                tableResponse.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                if (responseBody == null
                    || responseBody.Value.Value == null
                    || responseBody.Value.Value.Schema == null)
                {
                    continue;
                }

                int firstTableFieldsCellRow = currentRow;
                int firstTableFieldsCellColumn = 1;
                worksheet.Cell(currentRow, 1).Value = "Field name";
                worksheet.Cell(currentRow, 2).Value = "Field type";
                worksheet.Cell(currentRow, 3).Value = "Is mandatory?";
                currentRow++;
                // display simple fields owned by the current object
                foreach (var property in responseBody.Value.Value.Schema.Properties)
                {
                    var isPropertyRequired = responseBody.Value.Value.Schema.Required.Contains(property.Key);
                    GenerateProperty(worksheet, responseBody.Value.Value.Schema, property.Key, property.Value, ref currentRow, isPropertyRequired);
                }
                // display sub objects
                foreach (var schema in responseBody.Value.Value.Schema.AllOf)
                {
                    foreach (var property in schema.Properties)
                    {
                        var isPropertyRequired = schema.Required.Contains(property.Key);
                        GenerateProperty(worksheet, responseBody.Value.Value.Schema, property.Key, property.Value, ref currentRow, isPropertyRequired);
                    }
                }
                int oneOfIndex = 1;
                foreach (var schema in responseBody.Value.Value.Schema.OneOf)
                {
                    worksheet.Cell(currentRow, 1).Value = string.Format("Response can be of this schema #{0}", oneOfIndex);
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Font.Italic = true;
                    worksheet.Range(currentRow, 1, currentRow, 3).Merge();
                    currentRow++;
                    foreach (var property in schema.Properties)
                    {
                        var isPropertyRequired = schema.Required.Contains(property.Key);
                        GenerateProperty(worksheet, responseBody.Value.Value.Schema, property.Key, property.Value, ref currentRow, isPropertyRequired);
                    }
                    oneOfIndex++;
                }
                if (responseBody.Value.Value.Schema.Items != null)
                {

                    // this property node doesn't need to be displayed, so we will skip it
                    GenerateProperty(worksheet, responseBody.Value.Value.Schema, string.Empty, responseBody.Value.Value.Schema.Items, ref currentRow, true, true);
                }

                int lastTableFieldsCellRow = currentRow - 1; // -1 to position the cursor on the last row of the table
                int lastTableFieldsCellColumn = 3;

                // Apply table style
                var tableFields = worksheet.Range(firstTableFieldsCellRow, firstTableFieldsCellColumn, lastTableFieldsCellRow, lastTableFieldsCellColumn);
                var tableFieldsHeaders = tableFields.Range(1, 1, 1, tableFields.ColumnCount()); // coordonates are relative to the table
                tableFieldsHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                tableFieldsHeaders.Style.Font.Bold = true;
                tableFieldsHeaders.Style.Fill.BackgroundColor = XLColor.Orange;
                tableFields.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                tableFields.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                currentRow++; // add a new line to easy readability
            }

        }

        /// <summary>
        /// Generate information concerning a given property, and potentially its sub properties, if it is a complex type (array or object).
        /// </summary>
        /// <param name="worksheet">The worksheet where data have to be written</param>
        /// <param name="parentSchema">OpenApi schema of the parent object</param>
        /// <param name="propertyName">Name of the current property</param>
        /// <param name="propertySchema">Oas metadata about this property</param>
        /// <param name="currentRow">Current row of the worksheet, where data have to be written</param>
        /// <param name="isRequired">Indicates if the current property is a mandatory or optional field</param>
        /// <param name="skipProperty">In some cases, it is not relevant to write the property in the Excel file. This parameter enables to skip it, without skipping its potential sub properties (if any).</param>
        /// <param name="isLastPropertyLevel">Indicates if the current property has to be the last node or not. If yes, its child properties will not be displayed.</param>
        private void GenerateProperty(IXLWorksheet worksheet, OpenApiSchema parentSchema, string propertyName, OpenApiSchema propertySchema, ref int currentRow, bool isRequired = false, bool skipProperty = false, bool isLastPropertyLevel = false)
        {
            if (skipProperty == false)
            {
                worksheet.Cell(currentRow, 1).Value = propertyName;
                worksheet.Cell(currentRow, 2).Value = propertySchema.Type;
                worksheet.Cell(currentRow, 3).Value = isRequired;
                if (!string.IsNullOrEmpty(propertySchema.Description))
                {
                    worksheet.Cell(currentRow, 1).Comment.AddNewLine().AddText(propertySchema.Description);
                    worksheet.Cell(currentRow, 1).Comment.Style.Size.SetAutomaticSize();
                }
                currentRow++;
            }

            if(isLastPropertyLevel)
            {
                return;
            }


            foreach (var property in propertySchema.Properties)
            {
                bool isChildLastPropertyLevel = false;
                // this validation is required to stop object parsing with cycling references
                if (HasAtLeastOneSameSchema(parentSchema, property.Value))
                {
                    isChildLastPropertyLevel = true;
                }
                var isPropertyRequired = propertySchema.Required.Contains(property.Key);
                GenerateProperty(worksheet, propertySchema, string.Format("{0}.{1}", propertyName, property.Key), property.Value, ref currentRow, isPropertyRequired, false, isChildLastPropertyLevel);
            }
            foreach (var schema in propertySchema.AllOf)
            {
                foreach (var property in schema.Properties)
                {
                    bool isChildLastPropertyLevel = false;
                    // this validation is required to stop object parsing with cycling references
                    if (HasAtLeastOneSameSchema(parentSchema, property.Value))
                    {
                        isChildLastPropertyLevel = true;
                    }
                    var isPropertyRequired = schema.Required.Contains(property.Key);
                    GenerateProperty(worksheet, propertySchema, string.Format("{0}.{1}", propertyName, property.Key), property.Value, ref currentRow, isPropertyRequired, false, isChildLastPropertyLevel);
                }
            }
            int oneOfIndex = 1;
            foreach (var schema in propertySchema.OneOf)
            {
                worksheet.Cell(currentRow, 1).Value = string.Format("Sub object can be of this schema #{0}", oneOfIndex);
                worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                worksheet.Cell(currentRow, 1).Style.Font.Italic = true;
                worksheet.Range(currentRow, 1, currentRow, 3).Merge();
                currentRow++;
                foreach (var property in schema.Properties)
                {
                    bool isChildLastPropertyLevel = false;
                    // this validation is required to stop object parsing with cycling references
                    if (HasAtLeastOneSameSchema(parentSchema, property.Value))
                    {
                        isChildLastPropertyLevel = true;
                    }
                    var isPropertyRequired = schema.Required.Contains(property.Key);
                    GenerateProperty(worksheet, propertySchema, property.Key, property.Value, ref currentRow, isPropertyRequired, false, isChildLastPropertyLevel);
                }
                oneOfIndex++;
            }
            if (propertySchema.Items != null)
            {
                // this property node doesn't need to be displayed, so we skip it
                GenerateProperty(worksheet, propertySchema, propertyName, propertySchema.Items, ref currentRow, true, true, false);
            }

        }

        /// <summary>
        /// Compare two Open API schemas to identify if they have at least one identical sub schema (oneOf, allOf, anyOf, root schema)
        /// </summary>
        /// <param name="schema1">First schema to compare</param>
        /// <param name="schema2">Second schema to compare</param>
        /// <returns></returns>
        private bool HasAtLeastOneSameSchema(OpenApiSchema schema1, OpenApiSchema schema2)
        {
            List<OpenApiSchema> composedSchemas1 = new List<OpenApiSchema>();
            GetComposedSchemasRecursively(schema1, ref composedSchemas1);
            List<OpenApiSchema> composedSchemas2 = new List<OpenApiSchema>();
            GetComposedSchemasRecursively(schema2, ref composedSchemas2);

            foreach (OpenApiSchema sch in composedSchemas1)
            {
                if (composedSchemas2.Contains(sch))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Parse the root schema to extract recursively all sub schemas composing it.
        /// </summary>
        /// <param name="rootSchema"></param>
        /// <param name="composedSchemas"></param>
        private void GetComposedSchemasRecursively(OpenApiSchema rootSchema, ref List<OpenApiSchema> composedSchemas)
        {
            if (rootSchema == null)
            {
                return;
            }
            composedSchemas.Add(rootSchema);

            if (rootSchema.Items != null)
            {
                composedSchemas.Add(rootSchema.Items);
            }

            foreach (var allOfSchema in rootSchema.AllOf)
            {
                if (allOfSchema != null)
                {
                    composedSchemas.Add(allOfSchema);
                    GetComposedSchemasRecursively(allOfSchema, ref composedSchemas);
                }
            }
            foreach (var anyOfSchema in rootSchema.AnyOf)
            {
                if (anyOfSchema != null)
                {
                    composedSchemas.Add(anyOfSchema);
                    GetComposedSchemasRecursively(anyOfSchema, ref composedSchemas);
                }
            }
            foreach (var oneOfSchema in rootSchema.OneOf)
            {
                if (oneOfSchema != null)
                {
                    composedSchemas.Add(oneOfSchema);
                    GetComposedSchemasRecursively(oneOfSchema, ref composedSchemas);
                }
            }
        }

    }
}
