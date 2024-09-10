using System;
using System.IO;
using Newtonsoft.Json.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO.Compression;
using Microsoft.Data.SqlClient;
using System.Text;


namespace ParsedCVData
{
    class Program
    {
        private static readonly string _connectionString = "Server=DESKTOP-I41R0QQ;Database=TOLTournamentLeague;Trusted_Connection=True;TrustServerCertificate=True";

        //static async Task Main(string[] args)
        //{
        //    // Sample command-line arguments for demonstration
        //    if (args.Length > 0 && args[0] == "generate")
        //    {
        //        int limit = args.Length > 1 ? int.Parse(args[1]) : 10;
        //        await GenerateAndDownload(limit);
        //    }
        //    else if (args.Length > 0 && args[0] == "resume")
        //    {
        //        int id = int.Parse(args[1]);
        //        await GetResume(id);
        //    }
        //    else
        //    {
        //        Console.WriteLine("Invalid arguments. Use 'generate' or 'resume <id>'.");
        //    }
        //}
        static async Task Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("No arguments provided. Use 'generate' to generate PDFs or 'resume <id>' to retrieve a resume.");
                return;
            }

            if (args[0].Equals("generate", StringComparison.OrdinalIgnoreCase))
            {
                int limit = 10; // Default value
                if (args.Length > 1 && int.TryParse(args[1], out int parsedLimit))
                {
                    limit = parsedLimit;
                }
                await GenerateAndDownload(limit);
            }
            else if (args[0].Equals("resume", StringComparison.OrdinalIgnoreCase))
            {
                if (args.Length > 1 && int.TryParse(args[1], out int id))
                {
                    await GetResume(id);
                }
                else
                {
                    Console.WriteLine("Invalid or missing ID. Please provide a valid integer ID.");
                }
            }
            else
            {
                Console.WriteLine("Invalid arguments. Use 'generate' or 'resume <id>'.");
            }
        }


        private static async Task GetResume(int id)
        {
            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    await connection.OpenAsync();
                    var commandText = "SELECT ParsedData, isValidCV FROM tblParsedCVData WHERE Id = @Id;";
                    using (var cmd = new SqlCommand(commandText, connection))
                    {
                        cmd.Parameters.AddWithValue("@Id", id);
                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                var json = reader["ParsedData"] as string;
                                var isValidCV = Convert.ToBoolean(reader["isValidCV"]);

                                if (!isValidCV)
                                {
                                    Console.WriteLine("CV is not valid and cannot be generated.");
                                    return;
                                }

                                if (json == null)
                                {
                                    Console.WriteLine("No data found for the given ID.");
                                    return;
                                }

                                var jsonObject = JObject.Parse(json);
                                var candidateName = jsonObject["Name"]?.ToString() ?? "Candidate";
                                string imagePath = "C:\\Users\\Shoib\\Downloads\\talentonlease_cover__1removebg.png";
                                var pdfBytes = CreatePdfFromJson(json, imagePath);

                                var sanitizedCandidateName = System.Text.RegularExpressions.Regex.Replace(candidateName, @"[^a-zA-Z0-9_\-]", "_");
                                var fileName = $"{sanitizedCandidateName}_CV.pdf";

                                File.WriteAllBytes(fileName, pdfBytes);
                                Console.WriteLine($"PDF generated: {fileName}");
                            }
                            else
                            {
                                Console.WriteLine("No data found for the given ID.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving resume data: {ex.Message}");
            }
        }

        //private static async Task GenerateAndDownload(int limit)
        //{
        //    var dataModels = await GetValidDataAsync(limit);

        //    if (dataModels.Count == 0)
        //    {
        //        Console.WriteLine("No valid data found.");
        //        return;
        //    }


        //    string tempDir = Path.Combine(Path.GetTempPath(), "pdfs");
        //    Directory.CreateDirectory(tempDir);
        //    string zipFilePath = Path.Combine(@"C:\Users\Shoib\source\repos\ParsedCVData\IsStandardCVGenerate", "cv_files.zip");


        //    using (var zipToOpen = new FileStream(zipFilePath, FileMode.Create))
        //    using (var archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        //    {
        //        foreach (var dataModel in dataModels)
        //        {
        //            byte[] pdfBytes = CreatePdfFromJson(dataModel.ParsedData, "C:\\Users\\Shoib\\Downloads\\talentonlease_cover__1removebg.png");
        //            string candidateName = ExtractCandidateNameFromJson(dataModel.ParsedData);
        //            string fileName = ReplaceInvalidFileNameChars($"{candidateName}_{dataModel.Id}.pdf", '_');

        //            var zipEntry = archive.CreateEntry(fileName);
        //            using (var zipStream = zipEntry.Open())
        //            {
        //                zipStream.Write(pdfBytes, 0, pdfBytes.Length);
        //            }

        //            await UpdateStatusAsync(dataModel.Id);
        //        }
        //    }

        //    var fileBytes = await File.ReadAllBytesAsync(zipFilePath);
        //    //File.Delete(zipFilePath);
        //    //Directory.Delete(tempDir, true);

        //    Console.WriteLine("ZIP file created and cleaned up.");
        //}


        private static async Task GenerateAndDownload(int limit)
        {
            var dataModels = await GetValidDataAsync(limit);

            if (dataModels.Count == 0)
            {
                Console.WriteLine("No valid data found.");
                return;
            }

            // Define the directory and path for the ZIP file
            string targetDir = @"C:\Users\Shoib";
            Directory.CreateDirectory(targetDir);
            string zipFilePath = Path.Combine(targetDir, "cv_files.zip");

            using (var zipToOpen = new FileStream(zipFilePath, FileMode.Create))
            using (var archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
            {
                foreach (var dataModel in dataModels)
                {
                    byte[] pdfBytes = CreatePdfFromJson(dataModel.ParsedData, "C:\\Users\\Shoib\\Downloads\\talentonlease_cover__1removebg.png");
                    string candidateName = ExtractCandidateNameFromJson(dataModel.ParsedData);
                    string fileName = ReplaceInvalidFileNameChars($"{candidateName}_{dataModel.Id}.pdf", '_');

                    // Write PDF file to a temporary directory
                    string tempDir = Path.Combine(Path.GetTempPath(), "pdfs");
                    Directory.CreateDirectory(tempDir);
                    string pdfFilePath = Path.Combine(tempDir, fileName);
                    File.WriteAllBytes(pdfFilePath, pdfBytes);

                    // Add PDF to ZIP file
                    var zipEntry = archive.CreateEntry(fileName);
                    using (var zipStream = zipEntry.Open())
                    {
                        using (var pdfStream = new FileStream(pdfFilePath, FileMode.Open))
                        {
                            pdfStream.CopyTo(zipStream);
                        }
                    }

                    await UpdateStatusAsync(dataModel.Id);
                }
            }

            Console.WriteLine("ZIP file created and saved at: " + zipFilePath);
        }



        private static async Task<List<ParsedDatamodel>> GetValidDataAsync(int limit)
        {
            var result = new List<ParsedDatamodel>();

            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var query = @"
                SELECT TOP (@Limit) Id, ParsedData
                FROM tblParsedCVData
                WHERE IsStandardCVGenerate = 0
                AND IsValidCV = 1
                ORDER BY Id";

                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Limit", limit);

                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            result.Add(new ParsedDatamodel
                            {
                                Id = reader.GetInt32(0),
                                ParsedData = reader.GetString(1)
                            });
                        }
                    }
                }
            }

            return result;
        }

        private static string ExtractCandidateNameFromJson(string json)
        {
            try
            {
                var jsonObject = JObject.Parse(json);
                return jsonObject["Name"]?.ToString() ?? "UnknownCandidate";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing JSON: {ex.Message}");
                return "UnknownCandidate";
            }
        }

        private static string ReplaceInvalidFileNameChars(string fileName, char replacementChar)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, replacementChar);
            }
            return fileName;
        }

        private static async Task UpdateStatusAsync(int id)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var query = "UPDATE tblParsedCVData SET IsStandardCVGenerate = 1 WHERE Id = @Id";
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", id);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public static byte[] CreatePdfFromJson(string json, string imagePath)
        {
            imagePath = "C:\\Users\\Shoib\\Downloads\\talentonlease_cover__1removebg.png";

            using (MemoryStream ms = new MemoryStream())
            {
                Document doc = new Document(PageSize.A4);
                PdfWriter writer = PdfWriter.GetInstance(doc, ms);

                var pageEvents = new MyPageEvents(imagePath);
                writer.PageEvent = pageEvents;

                // Set page event to handle gray rectangle on every page
                writer.PageEvent = new GrayBackgroundPageEvent();

                try
                {
                    // Define colors
                    BaseColor titleColor = new BaseColor(27, 39, 76);
                    BaseColor sectionColor = BaseColor.BLACK;

                    doc.Open();

                    // Calculate the split point based on percentage
                    float splitPoint = doc.PageSize.Width * 0.2f;

                    // Convert JSON string to JObject for better manipulation
                    JObject cvData = JObject.Parse(json);

                    // Add Personal Information
                    AddPersonalInformation(doc, cvData, titleColor, sectionColor, splitPoint);

                    // Add Professional Summary
                    string professionalSummary = cvData["Professional Summary"]?.ToString();
                    if (!string.IsNullOrEmpty(professionalSummary))
                    {
                        AddSectionTitle(doc, "Professional Summary".ToUpper(), titleColor, sectionColor, splitPoint);
                        AddParagraph(doc, professionalSummary, sectionColor, splitPoint, justify: true);
                    }

                    // Add Skills
                    JToken skills = cvData["Skills"];
                    if (skills != null && skills.HasValues)
                    {
                        AddSectionTitle(doc, "Skills".ToUpper(), titleColor, sectionColor, splitPoint);
                        AddSkills(doc, skills, sectionColor, splitPoint);
                    }

                    // Add Education
                    JToken education = cvData["Education"];
                    if (education != null && education.HasValues)
                    {
                        AddSectionTitle(doc, "Education".ToUpper(), titleColor, sectionColor, splitPoint);
                        AddEducation(doc, education, sectionColor, splitPoint);
                    }

                    // Add Work Experience
                    JToken workExperience = cvData["Experience"];
                    if (workExperience != null && workExperience.HasValues)
                    {
                        AddSectionTitle(doc, "Work Experience".ToUpper(), titleColor, sectionColor, splitPoint);
                        AddWorkExperience(doc, workExperience, sectionColor, splitPoint);
                    }

                    // Determine which projects key to use
                    JToken projectData = cvData["ProjectDetails"] as JArray;
                    if (projectData == null)
                    {
                        projectData = cvData["Projects"] as JArray;
                    }
                    if (projectData == null)
                    {
                        projectData = cvData["Project_Details"] as JArray;
                    }

                    // Add Project Details
                    if (projectData != null && projectData.HasValues)
                    {
                        AddSectionTitle(doc, "Projects".ToUpper(), titleColor, sectionColor, splitPoint);
                        AddProjects(doc, projectData, sectionColor, splitPoint);
                    }

                    // Add Certifications
                    JToken certifications = cvData["Certifications"];
                    if (certifications != null && certifications.HasValues)
                    {
                        AddSectionTitle(doc, "Certifications".ToUpper(), titleColor, sectionColor, splitPoint);
                        AddCertifications(doc, certifications, sectionColor, splitPoint);
                    }
                }
                catch (DocumentException de)
                {
                    throw new IOException("Error creating PDF", de);
                }
                finally
                {
                    doc.Close();
                    writer.Close();
                }

                return ms.ToArray();
            }
        }

   
        private static void AddSkills(Document doc, JToken skillsToken, BaseColor sectionColor, float splitPoint)
        {
            if (skillsToken == null)
            {
                // If skillsToken is null, exit the method
                return;
            }

            if (skillsToken is JObject skillsObject)
            {
                // Handle JSON object format (categories with arrays of skills)
                foreach (var category in skillsObject)
                {
                    string categoryTitle = category.Key;
                    var skillsArray = category.Value as JArray;

                    if (skillsArray == null || !skillsArray.Any())
                        continue; // Skip empty or invalid skill arrays

                    // Join the skills array into a comma-separated string
                    string skillsString = string.Join(", ", skillsArray.Select(s => s.ToString()));

                    // Add Category Title
                    Paragraph categoryTitlePara = new Paragraph(categoryTitle, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, sectionColor))
                    {
                        Alignment = Element.ALIGN_LEFT,
                        IndentationLeft = splitPoint + 20f,
                        SpacingAfter = 5f
                    };
                    doc.Add(categoryTitlePara);

                    // Add Skills to the document
                    Paragraph skillsPara = new Paragraph(skillsString, FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
                    {
                        Alignment = Element.ALIGN_LEFT,
                        IndentationLeft = splitPoint + 20f,
                        SpacingAfter = 15f
                    };
                    doc.Add(skillsPara);
                }
            }
            else if (skillsToken is JArray skillsArray)
            {
                // Handle JSON array format
                string skillsString = string.Join(", ", skillsArray.Select(s => s.ToString()));

                // Add Skills to the document
                Paragraph skillsPara = new Paragraph(skillsString, FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
                {
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f,
                    SpacingAfter = 15f
                };
                doc.Add(skillsPara);
            }
        }

        private static void AddPersonalInformation(Document doc, JObject cvData, BaseColor titleColor, BaseColor sectionColor, float splitPoint)
        {
            // Add Name
            string fullName = cvData["Name"]?.ToString()?.ToUpper();
            if (!string.IsNullOrEmpty(fullName))
            {
                Paragraph namePara = new Paragraph(fullName, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 20, titleColor))
                {
                    Alignment = Element.ALIGN_LEFT,
                    SpacingAfter = 2f, // Add spacing after name
                    IndentationLeft = splitPoint + 20f
                };
                doc.Add(namePara);
            }

            // Add Job Title
            string jobTitle = cvData["Job Title"]?.ToString()?.ToUpper();
            if (!string.IsNullOrEmpty(jobTitle))
            {
                Paragraph jobTitlePara = new Paragraph(jobTitle, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, sectionColor))
                {
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f,
                    SpacingAfter = 15f
                };
                doc.Add(jobTitlePara);
            }

            // Initialize contact details
            string contactDetails = "";

            // Handle Phone Numbers
            JToken phoneToken = cvData["Phone Number"];
            if (phoneToken is JArray phoneArray && phoneArray.Any())
            {
                foreach (var phone in phoneArray.Where(phone => !string.IsNullOrWhiteSpace(phone.ToString())))
                {
                    contactDetails += $"\u2022 {phone.ToString()}\n"; // Bullet symbol •
                }
            }
            else if (phoneToken is JValue phoneValue && !string.IsNullOrWhiteSpace(phoneValue.ToString()))
            {
                contactDetails += $"\u2022 {phoneValue.ToString()}\n"; // Bullet symbol •
            }

            // Handle Emails
            JToken emailToken = cvData["Email"];
            if (emailToken is JArray emailArray && emailArray.Any())
            {
                foreach (var email in emailArray.Where(email => !string.IsNullOrWhiteSpace(email.ToString())))
                {
                    contactDetails += $"\u2022 {email.ToString()}\n"; // Bullet symbol •
                }
            }
            else if (emailToken is JValue emailValue && !string.IsNullOrWhiteSpace(emailValue.ToString()))
            {
                contactDetails += $"\u2022 {emailValue.ToString()}\n"; // Bullet symbol •
            }

            // If no contact details were added, include a placeholder
            if (string.IsNullOrEmpty(contactDetails))
            {
                
            }

            // Add Contact Information to the document
            Paragraph contactPara = new Paragraph(contactDetails, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, sectionColor))
            {
                Alignment = Element.ALIGN_LEFT,
                IndentationLeft = splitPoint + 20f,
                SpacingAfter = 20f
            };
            doc.Add(contactPara);

        }
        private static void AddWorkExperience(Document doc, JToken workExperience, BaseColor sectionColor, float splitPoint)
        {
            if (workExperience != null && workExperience is JArray workExperienceArray && workExperienceArray.Any())
            {
                foreach (var item in workExperienceArray)
                {
                    // Extract job title, company name, and dates/date
                    string jobTitle = GetValueOrNull(item, "Job Title")?.ToUpper();
                    string companyName = GetValueOrNull(item, "Company")?.ToUpper();
                    string dates = GetValueOrNull(item, "Dates") ?? GetValueOrNull(item, "Date");

                    if (!string.IsNullOrEmpty(jobTitle) && !string.IsNullOrEmpty(companyName))
                    {
                        // Add job title and company name (Bold)
                        Paragraph headerPara = new Paragraph($"{jobTitle} at {companyName}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, new BaseColor(27, 39, 76)))
                        {
                            SpacingAfter = 5f,
                            Alignment = Element.ALIGN_LEFT,
                            IndentationLeft = splitPoint + 20f
                        };
                        doc.Add(headerPara);

                        // Add dates or date
                        if (!string.IsNullOrEmpty(dates))
                        {
                            Paragraph datesPara = new Paragraph($"{dates}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLUE))
                            {
                                SpacingAfter = 5f,
                                Alignment = Element.ALIGN_LEFT,
                                IndentationLeft = splitPoint + 20f
                            };
                            doc.Add(datesPara);
                        }

                        // Get responsibilities as an array of strings if available
                        JArray responsibilitiesArray = item["Responsibilities"] as JArray;
                        List<string> responsibilitiesList = new List<string>();

                        if (responsibilitiesArray != null && responsibilitiesArray.Any())
                        {
                            responsibilitiesList.AddRange(responsibilitiesArray.Select(r => r.ToString()));
                        }
                        else
                        {
                            // Check if responsibilities are provided as a single string
                            string responsibilitiesString = GetValueOrNull(item, "Responsibilities");
                            if (!string.IsNullOrEmpty(responsibilitiesString))
                            {
                                responsibilitiesList.Add(responsibilitiesString);
                            }
                        }

                        // Create a bullet point list for responsibilities with circular bullets
                        if (responsibilitiesList.Any())
                        {
                            List responsibilitiesListItem = new List(List.UNORDERED)
                            {
                                IndentationLeft = splitPoint + 20f,
                                Symbol = new Chunk("• ", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor)) // This sets the bullet symbol
                            };

                            foreach (var responsibility in responsibilitiesList)
                            {
                                responsibilitiesListItem.Add(new ListItem(responsibility, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor)));
                            }

                            // Add responsibilities list
                            doc.Add(responsibilitiesListItem);
                        }
                        else
                        {
                            // Optionally add a placeholder if there are no responsibilities
                            doc.Add(new Paragraph("No responsibilities listed.", FontFactory.GetFont(FontFactory.HELVETICA, 10, sectionColor)));
                        }

                        // Add additional space between entries
                        doc.Add(new Paragraph("\n")); // Adds a line break
                    }
                }
            }
            else
            {
                // If no work experience is available
                // Skip adding the section
            }
        }

        private static void AddCertifications(Document doc, JToken certifications, BaseColor sectionColor, float splitPoint)
        {
            if (certifications != null && certifications is JArray certificationsArray && certificationsArray.Any())
            {
                foreach (var cert in certificationsArray)
                {
                    // Retrieve and format each field
                    string certName = GetValueOrNull(cert, "Certification Name")?.ToUpper();
                    string issuingOrg = GetValueOrNull(cert, "Issuing Organization")?.ToUpper();
                    string certDate = GetValueOrNull(cert, "Date");

                    if (!string.IsNullOrEmpty(certName))
                    {
                        // Add Certification Name
                        Paragraph certNamePara = new Paragraph(certName, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, new BaseColor(27, 39, 76)))
                        {
                            Alignment = Element.ALIGN_LEFT,
                            IndentationLeft = splitPoint + 20f,
                            SpacingAfter = 2f
                        };
                        doc.Add(certNamePara);

                        // Add Issuing Organization
                        if (!string.IsNullOrEmpty(issuingOrg))
                        {
                            Paragraph issuingOrgPara = new Paragraph($"Issuing Organization: {issuingOrg}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor))
                            {
                                Alignment = Element.ALIGN_LEFT,
                                IndentationLeft = splitPoint + 20f,
                                SpacingAfter = 2f
                            };
                            doc.Add(issuingOrgPara);
                        }

                        // Add Date
                        if (!string.IsNullOrEmpty(certDate))
                        {
                            Paragraph certDatePara = new Paragraph($"Date: {certDate}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLUE))
                            {
                                Alignment = Element.ALIGN_LEFT,
                                IndentationLeft = splitPoint + 20f,
                                SpacingAfter = 10f
                            };
                            doc.Add(certDatePara);
                        }

                        // Add a line break between certifications
                        doc.Add(new Paragraph("\n"));
                    }
                }
            }
            else
            {
                // If no certifications are available
                // Skip adding the section
            }
        }

     
        private static void AddProjects(Document doc, JToken projects, BaseColor sectionColor, float splitPoint)
        {
            if (projects == null)
            {
                return; // Skip adding projects section
            }

            // Determine if projects is an array or an object
            JArray projectsArray = null;

            if (projects is JObject obj)
            {
                // Check for "Project Details" first
                if (obj["Project Details"] is JArray projectDetailsArray && projectDetailsArray.Any())
                {
                    projectsArray = projectDetailsArray;
                }
                // Check for "Projects" next
                else if (obj["Projects"] is JArray projectsArrayFromObject && projectsArrayFromObject.Any())
                {
                    projectsArray = projectsArrayFromObject;
                }
                // Check for "Project_Details" as the last fallback
                else if (obj["Project_Details"] is JArray projectDetailsArray2 && projectDetailsArray2.Any())
                {
                    projectsArray = projectDetailsArray2;
                }
                else
                {
                    Console.WriteLine("No valid project data found.");
                    return;
                }
            }
            else if (projects is JArray array && array.Any())
            {
                projectsArray = array;
            }
            else
            {
                Console.WriteLine("Projects data is not an array or object.");
                return;
            }

            // Loop through each project item
            foreach (var item in projectsArray)
            {
                if (item is not JObject projectItem)
                {
                    Console.WriteLine("Project item is not a valid object.");
                    continue;
                }

                // Extracting fields from the item
                string projectName = GetValueOrNull(projectItem, "Project Name")?.ToUpper();
                string projectName2 = GetValueOrNull(projectItem, "Project_Name")?.ToUpper(); // Added for "Project_Name"
                string projectToDisplay = !string.IsNullOrEmpty(projectName) ? projectName : projectName2;

                // If both projectName and projectName2 are null or empty, skip this item
                if (string.IsNullOrEmpty(projectToDisplay))
                {
                    Console.WriteLine("Project name is not available for the current item.");
                    continue;
                }

                string description = GetValueOrNull(projectItem, "Description")?.ToLower() ?? "N/A";

                // Initialize technologies and roles as empty strings
                string technologies = "";
                string roles = "";

                // Handle Technologies Used
                JToken technologiesToken = projectItem["Technologies Used"];
                if (technologiesToken is JValue technologiesValue)
                {
                    technologies = technologiesValue.ToString();
                }
                else if (technologiesToken is JArray technologiesArray && technologiesArray.Any())
                {
                    technologies = string.Join(", ", technologiesArray.Select(t => t.ToString()));
                }

                // Handle Roles (both "Role" and "Roles")
                List<string> rolesList = new List<string>();

                JToken rolesToken = projectItem["Role"];
                if (rolesToken is JValue rolesValue)
                {
                    rolesList.Add(rolesValue.ToString());
                }
                else if (rolesToken is JArray rolesArray && rolesArray.Any())
                {
                    rolesList.AddRange(rolesArray.Select(r => r.ToString()));
                }

                JToken rolesTokenAlt = projectItem["Roles"];
                if (rolesTokenAlt is JValue rolesValueAlt)
                {
                    rolesList.Add(rolesValueAlt.ToString());
                }
                else if (rolesTokenAlt is JArray rolesArrayAlt && rolesArrayAlt.Any())
                {
                    rolesList.AddRange(rolesArrayAlt.Select(r => r.ToString()));
                }

                // Join rolesList into a comma-separated string
                roles = string.Join(", ", rolesList);

                // Add Project Name (Bold)
                Paragraph projectHeaderPara = new Paragraph($"{projectToDisplay}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, new BaseColor(27, 39, 76)))
                {
                    SpacingAfter = 5f,
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f
                };
                doc.Add(projectHeaderPara);

                // Add Project Description, Technologies, and Roles
                StringBuilder detailsBuilder = new StringBuilder();
                detailsBuilder.AppendLine($"Description: {description}");
                if (!string.IsNullOrEmpty(technologies))
                {
                    detailsBuilder.AppendLine($"Technologies: {technologies}");
                }
                if (!string.IsNullOrEmpty(roles))
                {
                    detailsBuilder.AppendLine($"Role: {roles}");
                }

                Paragraph projectDetailsPara = new Paragraph(detailsBuilder.ToString(), FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor))
                {
                    SpacingAfter = 10f,
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f
                };
                doc.Add(projectDetailsPara);

                // Add additional space between projects
                doc.Add(new Paragraph("\n")); // Adds a line break
            }
        }


        private static void AddEducation(Document doc, JToken education, BaseColor sectionColor, float splitPoint)
        {
            if (education != null)
            {
                if (education is JArray educationArray && educationArray.Any())
                {
                    // Handle the case where education is an array of objects
                    foreach (var item in educationArray)
                    {
                        string degree = item["Degree"]?.ToString();
                        string institution = item["Institution"]?.ToString();
                        string graduationDate = item["Graduation Date"]?.ToString();

                        if (!string.IsNullOrEmpty(degree) || !string.IsNullOrEmpty(institution) || !string.IsNullOrEmpty(graduationDate))
                        {
                            // Build the paragraph with bold degree and normal text for the rest
                            Paragraph eduPara = new Paragraph
                            {
                                SpacingAfter = 10f,
                                Alignment = Element.ALIGN_LEFT,
                                IndentationLeft = splitPoint + 20f
                            };

                            // Add Degree as bold text
                            if (!string.IsNullOrEmpty(degree))
                            {
                                Chunk degreeChunk = new Chunk($"{degree}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, new BaseColor(27, 39, 76)));
                                eduPara.Add(degreeChunk);
                                eduPara.Add(new Chunk("\n")); // Add newline
                            }

                            // Add Institution as normal text
                            if (!string.IsNullOrEmpty(institution))
                            {
                                Chunk institutionChunk = new Chunk($"Institution: {institution}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor));
                                eduPara.Add(institutionChunk);
                                eduPara.Add(new Chunk("\n")); // Add newline
                            }

                            // Add Graduation Date as normal text
                            if (!string.IsNullOrEmpty(graduationDate))
                            {
                                Chunk graduationDateChunk = new Chunk($"Graduation Date: {graduationDate}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLUE));
                                eduPara.Add(graduationDateChunk);
                            }

                            doc.Add(eduPara);
                        }
                    }
                }
                else if (education is JObject educationObject)
                {
                    // Handle the case where education is a single object
                    string degree = educationObject["Degree"]?.ToString();
                    string institution = educationObject["Institution"]?.ToString();
                    string graduationDate = educationObject["Graduation Date"]?.ToString(); // Updated key for single object format

                    if (!string.IsNullOrEmpty(degree) || !string.IsNullOrEmpty(institution) || !string.IsNullOrEmpty(graduationDate))
                    {
                        // Build the paragraph with bold degree and normal text for the rest
                        Paragraph eduPara = new Paragraph
                        {
                            SpacingAfter = 10f,
                            Alignment = Element.ALIGN_LEFT,
                            IndentationLeft = splitPoint + 20f
                        };

                        // Add Degree as bold text
                        if (!string.IsNullOrEmpty(degree))
                        {
                            Chunk degreeChunk = new Chunk($"{degree}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, sectionColor));
                            eduPara.Add(degreeChunk);
                            eduPara.Add(new Chunk("\n")); // Add newline
                        }

                        // Add Institution as normal text
                        if (!string.IsNullOrEmpty(institution))
                        {
                            Chunk institutionChunk = new Chunk($"Institution: {institution}", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor));
                            eduPara.Add(institutionChunk);
                            eduPara.Add(new Chunk("\n")); // Add newline
                        }

                        // Add Graduation Date as normal text
                        if (!string.IsNullOrEmpty(graduationDate))
                        {
                            Chunk graduationDateChunk = new Chunk($"Graduation Date: {graduationDate}", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor));
                            eduPara.Add(graduationDateChunk);
                        }

                        doc.Add(eduPara);
                    }
                }
                else
                {
                    // If education is neither an array nor an object
                    return; // Skip adding the education section
                }
            }
            else
            {
                // If no education data is available
                return; // Skip adding the education section
            }
        }

        private static string GetValueOrNull(JToken token, string propertyName)
        {
            JToken value = token?[propertyName];
            return value?.ToString();
        }

        private static void AddSectionTitle(Document doc, string title, BaseColor titleColor, BaseColor sectionColor, float splitPoint)
        {
            Paragraph titlePara = new Paragraph(title, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, titleColor))
            {
                Alignment = Element.ALIGN_LEFT,
                SpacingBefore = 20f,
                SpacingAfter = 10f,
                IndentationLeft = splitPoint + 20f
            };
            doc.Add(titlePara);
        }

        private static void AddParagraph(Document doc, string text, BaseColor sectionColor, float splitPoint, bool justify = false)
        {
            Paragraph para = new Paragraph(text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, sectionColor))
            {
                Alignment = justify ? Element.ALIGN_JUSTIFIED : Element.ALIGN_LEFT,
                IndentationLeft = splitPoint + 20f,
                SpacingAfter = 15f
            };
            doc.Add(para);
        }

        private static void AddParagraphLeftAligned(Document doc, string text, BaseColor sectionColor, float splitPoint, bool bold = false)
        {
            // Create a font with optional bold style
            Font font = bold
                ? FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, sectionColor)
                : FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor);

            Paragraph para = new Paragraph(text, font)
            {
                Alignment = Element.ALIGN_LEFT,
                IndentationLeft = splitPoint + 20f,
                SpacingAfter = 15f
            };
            doc.Add(para);
        }



        public class GrayBackgroundPageEvent : PdfPageEventHelper
        {
            private readonly BaseColor pageColor = new BaseColor(255, 255, 255); // White color for the entire page
            private readonly string _footerText = "Copyright TalentOnLease © 2023. All Rights Reserved."; // Text to display in the footer
            private readonly float footerHeight = 30f; // Height of the footer

            private readonly BaseColor yellowColor = new BaseColor(27, 39, 76); // Color for the left portion
            private readonly BaseColor secondColor = new BaseColor(228, 44, 52); // Second color next to yellow

            public override void OnStartPage(PdfWriter writer, Document document)
            {
                PdfContentByte cb = writer.DirectContentUnder;
                // Fill the entire page with white color
                cb.SetColorFill(pageColor);
                cb.Rectangle(0, 0, document.PageSize.Width, document.PageSize.Height);
                cb.Fill();

                // Fill the left 10% of the page with the first color
                float yellowPortionWidth = document.PageSize.Width * 0.1f;
                cb.SetColorFill(yellowColor);
                cb.Rectangle(0, 0, yellowPortionWidth, document.PageSize.Height);
                cb.Fill();

                // Fill the next 10% of the page with the second color
                float secondColorPortionWidth = document.PageSize.Width * 0.1f;
                cb.SetColorFill(secondColor);
                cb.Rectangle(yellowPortionWidth, 0, secondColorPortionWidth, document.PageSize.Height);
                cb.Fill();
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {
                PdfContentByte canvas = writer.DirectContent;
                Font footerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 5, BaseColor.LIGHT_GRAY);
                float pageWidth = document.PageSize.Width;
                float footerYPosition = document.BottomMargin - 30f;

                Paragraph footerPara = new Paragraph(_footerText, footerFont)
                {
                    Alignment = Element.ALIGN_LEFT,
                    SpacingBefore = 1f
                };

                ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase(footerPara), pageWidth / 2, footerYPosition, 0);
            }
          
        }

        public class MyPageEvents : PdfPageEventHelper
        {
            private readonly string _imagePath;

            public MyPageEvents(string imagePath)
            {
                _imagePath = imagePath;
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {
                // Load the image
                iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(_imagePath);

                // Scale the image to fit smaller dimensions, e.g., 100x100
                img.ScaleToFit(100f, 100f);

                // Get the page dimensions
                float pageWidth = document.PageSize.Width;
                float pageHeight = document.PageSize.Height;

                // Calculate image position to be at the top-right corner
                float xPosition = pageWidth - img.ScaledWidth - 20f; // 20f margin from the right
                float yPosition = pageHeight - img.ScaledHeight - 20f; // 20f margin from the top

                // Set image position
                img.SetAbsolutePosition(xPosition, yPosition);

                // Add image to document
                writer.DirectContent.AddImage(img);
            }
        }

        public class ParsedDatamodel
        {
            public int Id { get; set; }
            public string ParsedData { get; set; }
        }
    }
}
