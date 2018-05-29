using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ChromeApplication.Startup))]
namespace ChromeApplication
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
