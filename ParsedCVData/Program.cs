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
            string targetDir = @"C:\Users\Shoib\source\repos\ParsedCVData\IsStandardCVGenerate";
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
                    AddSectionTitle(doc, "Professional Summary".ToUpper(), titleColor, sectionColor, splitPoint);
                    AddParagraph(doc, cvData["Professional Summary"]?.ToString() ?? "N/A", sectionColor, splitPoint, justify: true);

                    // Add Skills
                    AddSectionTitle(doc, "Skills".ToUpper(), titleColor, sectionColor, splitPoint);
                    AddSkills(doc, cvData["Skills"], sectionColor, splitPoint);

                    // Add Education
                    AddSectionTitle(doc, "Education".ToUpper(), titleColor, sectionColor, splitPoint);
                    AddEducation(doc, cvData["Education"], sectionColor, splitPoint);

                    // Add Work Experience
                    AddSectionTitle(doc, "Work Experience".ToUpper(), titleColor, sectionColor, splitPoint);
                    AddWorkExperience(doc, cvData["Experience"], sectionColor, splitPoint);

                    // Determine which projects key to use
                    JToken projectData = null;

                    if (cvData["Project Details"] is JArray projectDetailsArray && projectDetailsArray.Any())
                    {
                        projectData = projectDetailsArray;
                    }
                    else if (cvData["Projects"] is JArray projectsArray && projectsArray.Any())
                    {
                        projectData = projectsArray;
                    }
                    else if (cvData["Project_Details"] is JArray projectDetailsArray2 && projectDetailsArray2.Any())
                    {
                        projectData = projectDetailsArray2;
                    }
                    else
                    {
                        Console.WriteLine("No valid project data found.");
                    }

                    // Add Project Details
                    AddSectionTitle(doc, "Projects".ToUpper(), titleColor, sectionColor, splitPoint);
                    AddProjects(doc, projectData, sectionColor, splitPoint);

                    // Add Certifications
                    AddSectionTitle(doc, "Certifications".ToUpper(), titleColor, sectionColor, splitPoint);
                    AddCertifications(doc, cvData["Certifications"], sectionColor, splitPoint);
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
            // Check if skillsToken is null or not an array
            if (skillsToken == null || !(skillsToken is JArray skillsArray))
            {
                // If no skills available or not in the expected format
                Paragraph skillsPara = new Paragraph("Skills details not available", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
                {
                    SpacingAfter = 10f,
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f
                };
                doc.Add(skillsPara);
                return;
            }

            // Join the skills array into a comma-separated string
            string skillsString = string.Join(", ", skillsArray.Select(s => s.ToString()));

            // Add Skills to the document
            Paragraph skillsParaLeftAligned = new Paragraph(skillsString, FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
            {
                Alignment = Element.ALIGN_LEFT,
                IndentationLeft = splitPoint + 20f,
                SpacingAfter = 15f
            };
            doc.Add(skillsParaLeftAligned);
        }
        private static void AddPersonalInformation(Document doc, JObject cvData, BaseColor titleColor, BaseColor sectionColor, float splitPoint)
        {
            // Add Name
            string fullName = cvData["Name"]?.ToString()?.ToUpper() ?? "N/A";
            Paragraph namePara = new Paragraph(fullName, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 20, titleColor))
            {
                Alignment = Element.ALIGN_LEFT,
                SpacingAfter = 2f, // Add spacing after name
                IndentationLeft = splitPoint + 20f
            };
            doc.Add(namePara);

            // Add Job Title
            string jobTitle = cvData["Job Title"]?.ToString()?.ToUpper() ?? "N/A";
            Paragraph jobTitlePara = new Paragraph(jobTitle, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, sectionColor))
            {
                Alignment = Element.ALIGN_LEFT,
                IndentationLeft = splitPoint + 20f,
                SpacingAfter = 15f
            };
            doc.Add(jobTitlePara);

            // Initialize contact details
            string contactDetails = "";

            // Handle Phone Numbers
            JToken phoneToken = cvData["Phone Number"];
            if (phoneToken is JArray phoneArray)
            {
                foreach (var phone in phoneArray)
                {
                    contactDetails += $"\u2022 {phone.ToString()}\n"; // Bullet symbol •
                }
            }
            else if (phoneToken is JValue phoneValue)
            {
                contactDetails += $"\u2022 {phoneValue.ToString()}\n"; // Bullet symbol •
            }

            // Handle Emails
            JToken emailToken = cvData["Email"];
            if (emailToken is JArray emailArray)
            {
                foreach (var email in emailArray)
                {
                    contactDetails += $"\u2022 {email.ToString()}\n"; // Bullet symbol •
                }
            }
            else if (emailToken is JValue emailValue)
            {
                contactDetails += $"\u2022 {emailValue.ToString()}\n"; // Bullet symbol •
            }

            // If no contact details were added, include a placeholder
            if (string.IsNullOrEmpty(contactDetails))
            {
                contactDetails = "Contact details not available\n";
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
            if (workExperience != null && workExperience is JArray workExperienceArray)
            {
                foreach (var item in workExperienceArray)
                {
                    // Extract job title, company name, and dates/date
                    string jobTitle = GetValueOrNull(item, "Job Title")?.ToUpper() ?? "N/A";
                    string companyName = GetValueOrNull(item, "Company")?.ToUpper() ?? "N/A";

                    // Check for both "Dates" and "Date" fields
                    string dates = GetValueOrNull(item, "Dates") ?? GetValueOrNull(item, "Date") ?? "N/A";

                    // Get responsibilities as an array of strings if available
                    JArray responsibilitiesArray = item["Responsibilities"] as JArray;
                    List<string> responsibilitiesList = new List<string>();

                    if (responsibilitiesArray != null)
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

                    // Check if job title and company name are not empty before adding to the document
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
                        Paragraph datesPara = new Paragraph($"{dates}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLUE))
                        {
                            SpacingAfter = 5f,
                            Alignment = Element.ALIGN_LEFT,
                            IndentationLeft = splitPoint + 20f
                        };
                        doc.Add(datesPara);

                        // Create a bullet point list for responsibilities with circular bullets
                        List responsibilitiesListItem = new List(List.UNORDERED)
                        {
                            IndentationLeft = splitPoint + 20f,
                            Symbol = new Chunk("• ", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor)) // This sets the bullet symbol
                        };

                        if (responsibilitiesList.Any())
                        {
                            foreach (var responsibility in responsibilitiesList)
                            {
                                responsibilitiesListItem.Add(new ListItem(responsibility, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor)));
                            }
                        }
                        else
                        {
                            // Optionally add a placeholder if there are no responsibilities
                            responsibilitiesListItem.Add(new ListItem("N/A", FontFactory.GetFont(FontFactory.HELVETICA, 10, sectionColor)));
                        }

                        // Add responsibilities list
                        doc.Add(responsibilitiesListItem);

                        // Add additional space between entries
                        doc.Add(new Paragraph("\n")); // Adds a line break
                    }
                }
            }
            else
            {
                // If no work experience is available
                Paragraph expPara = new Paragraph("Work experience details not available", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
                {
                    SpacingAfter = 10f,
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f
                };
                doc.Add(expPara);
            }
        }
        private static void AddCertifications(Document doc, JToken certifications, BaseColor sectionColor, float splitPoint)
        {
            if (certifications != null && certifications is JArray certificationsArray)
            {
                foreach (var cert in certificationsArray)
                {
                    // Retrieve and format each field
                    string certName = GetValueOrNull(cert, "Certification Name")?.ToUpper() ?? "N/A";
                    string issuingOrg = GetValueOrNull(cert, "Issuing Organization")?.ToUpper() ?? "N/A";
                    string certDate = GetValueOrNull(cert, "Date") ?? "N/A";

                    // Add Certification Name
                    Paragraph certNamePara = new Paragraph(certName, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, new BaseColor(27, 39, 76)))
                    {
                        Alignment = Element.ALIGN_LEFT,
                        IndentationLeft = splitPoint + 20f,
                        SpacingAfter = 2f
                    };
                    doc.Add(certNamePara);

                    // Add Issuing Organization
                    Paragraph issuingOrgPara = new Paragraph($"Issuing Organization: {issuingOrg}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor))
                    {
                        Alignment = Element.ALIGN_LEFT,
                        IndentationLeft = splitPoint + 20f,
                        SpacingAfter = 2f
                    };
                    doc.Add(issuingOrgPara);

                    // Add Date
                    Paragraph certDatePara = new Paragraph($"Date: {certDate}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLUE))
                    {
                        Alignment = Element.ALIGN_LEFT,
                        IndentationLeft = splitPoint + 20f,
                        SpacingAfter = 10f
                    };
                    doc.Add(certDatePara);

                    // Add a line break between certifications
                    doc.Add(new Paragraph("\n"));
                }
            }
            else
            {
                // If no certifications are available
                Paragraph noCertPara = new Paragraph("Certification details not available", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
                {
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f,
                    SpacingAfter = 10f
                };
                doc.Add(noCertPara);
            }
        }
        private static void AddProjects(Document doc, JToken projects, BaseColor sectionColor, float splitPoint)
        {
            // Check if projects is null
            if (projects == null)
            {
                Paragraph projectPara = new Paragraph("No data available", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
                {
                    SpacingAfter = 15f,
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f
                };
                doc.Add(projectPara);
                return;
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
                // Handle cases where item might not be an object
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
                else if (technologiesToken is JArray technologiesArray)
                {
                    technologies = string.Join(", ", technologiesArray.Select(t => t.ToString()));
                }

                // Handle Roles
                JToken rolesToken = projectItem["Role"];
                if (rolesToken is JValue rolesValue)
                {
                    roles = rolesValue.ToString();
                }
                else if (rolesToken is JArray rolesArray)
                {
                    roles = string.Join(", ", rolesArray.Select(r => r.ToString()));
                }

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
            }
        }

        private static void AddEducation(Document doc, JToken education, BaseColor sectionColor, float splitPoint)
        {
            if (education != null)
            {
                if (education is JArray educationArray)
                {
                    // Handle the case where education is an array of objects
                    foreach (var item in educationArray)
                    {
                        string degree = item["Degree"]?.ToString() ?? "N/A";
                        string institution = item["Institution"]?.ToString() ?? "N/A";
                        string graduationDate = item["Graduation Date"]?.ToString() ?? "N/A";

                        // Build the paragraph with bold degree and normal text for the rest
                        Paragraph eduPara = new Paragraph
                        {
                            SpacingAfter = 10f,
                            Alignment = Element.ALIGN_LEFT,
                            IndentationLeft = splitPoint + 20f
                        };

                        // Add Degree as bold text
                        Chunk degreeChunk = new Chunk($"{degree}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, new BaseColor(27, 39, 76)));
                        eduPara.Add(degreeChunk);
                        eduPara.Add(new Chunk("\n")); // Add newline

                        // Add Institution as normal text
                        Chunk institutionChunk = new Chunk($"Institution: {institution}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, sectionColor));
                        eduPara.Add(institutionChunk);
                        eduPara.Add(new Chunk("\n")); // Add newline

                        // Add Graduation Date as normal text
                        Chunk graduationDateChunk = new Chunk($"Graduation Date: {graduationDate}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLUE));
                        eduPara.Add(graduationDateChunk);

                        doc.Add(eduPara);
                    }
                }
                else if (education is JObject educationObject)
                {
                    // Handle the case where education is a single object
                    string degree = educationObject["Degree"]?.ToString() ?? "N/A";
                    string institution = educationObject["Institution"]?.ToString() ?? "N/A";
                    string graduationDate = educationObject["Graduation Date"]?.ToString() ?? "N/A"; // Updated key for single object format

                    // Build the paragraph with bold degree and normal text for the rest
                    Paragraph eduPara = new Paragraph
                    {
                        SpacingAfter = 10f,
                        Alignment = Element.ALIGN_LEFT,
                        IndentationLeft = splitPoint + 20f
                    };

                    // Add Degree as bold text
                    Chunk degreeChunk = new Chunk($"{degree}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, sectionColor));
                    eduPara.Add(degreeChunk);
                    eduPara.Add(new Chunk("\n")); // Add newline

                    // Add Institution as normal text
                    Chunk institutionChunk = new Chunk($"Institution: {institution}", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor));
                    eduPara.Add(institutionChunk);
                    eduPara.Add(new Chunk("\n")); // Add newline

                    // Add Graduation Date as normal text
                    Chunk graduationDateChunk = new Chunk($"Graduation Date: {graduationDate}", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor));
                    eduPara.Add(graduationDateChunk);

                    doc.Add(eduPara);
                }
                else
                {
                    // If education is neither an array nor an object
                    Paragraph eduPara = new Paragraph("Education details not available", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
                    {
                        SpacingAfter = 10f,
                        Alignment = Element.ALIGN_LEFT,
                        IndentationLeft = splitPoint + 20f
                    };
                    doc.Add(eduPara);
                }
            }
            else
            {
                // If no education data is available
                Paragraph eduPara = new Paragraph("Education details not available", FontFactory.GetFont(FontFactory.HELVETICA, 12, sectionColor))
                {
                    SpacingAfter = 10f,
                    Alignment = Element.ALIGN_LEFT,
                    IndentationLeft = splitPoint + 20f
                };
                doc.Add(eduPara);
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
