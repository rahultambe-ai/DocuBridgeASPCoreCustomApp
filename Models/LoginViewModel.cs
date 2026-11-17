using System.ComponentModel.DataAnnotations;

namespace SSL_Core_Web_App_Identify_Sample.Models;

public class LoginViewModel
{
    [Required]
    [Display(Name = "User name")]
    public string UserName { get; set; } = "";

    [Required]
    [DataType(DataType.Password)]
    public string Password { get; set; } = "";

    [Display(Name = "Keep me signed in")]
    public bool RememberMe { get; set; }

    public string? ReturnUrl { get; set; }
}
