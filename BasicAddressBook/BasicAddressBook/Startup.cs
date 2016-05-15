using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(HanyAddressBookGet.Startup))]
namespace HanyAddressBookGet
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
