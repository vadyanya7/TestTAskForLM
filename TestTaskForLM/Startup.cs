using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(TestTaskForLM.Startup))]
namespace TestTaskForLM
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
