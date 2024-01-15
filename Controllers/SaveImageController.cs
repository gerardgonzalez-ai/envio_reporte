using Microsoft.AspNetCore.Mvc;
using EnvioInforme;
using Microsoft.Graph;

namespace SaveImage.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class SaveImageController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<SaveImageController> _logger;
        private IWebHostEnvironment _environment;

        public SaveImageController(IConfiguration config, ILogger<SaveImageController> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _config = config;
            _environment = environment;
        }

        [HttpPost(Name = "SaveImage")]
        public async Task<OperationResponse> SaveImage(ImageRequest request)
        {
            OperationResponse result = new OperationResponse();

            try
            {
                //TODO: Save image on local folder

                string rootPath = _environment.WebRootPath;
                string folderPath = Path.Combine(rootPath, "Images");
                bool isDirectoryExists = System.IO.Directory.Exists(folderPath);
                if (!isDirectoryExists)
                {
                    System.IO.Directory.CreateDirectory(folderPath);
                }

                byte[] imageBytes = Convert.FromBase64String(request.ImageB64);
                //System.IO.File.Create(folderPath + request.ImageName);
                System.IO.File.WriteAllBytes(Path.Combine(folderPath, request.ImageName), imageBytes);

                result.OperationResult = true;

                return result;
            }
            catch (Exception)
            {
                result.OperationResult = false;
                return result;
            }
        }

    }
}