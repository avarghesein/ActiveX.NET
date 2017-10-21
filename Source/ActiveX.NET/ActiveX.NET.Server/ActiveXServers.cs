using ActiveX.NET.Common;
using ActiveX.NET.Server.Utility;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Server
{
    internal class ActiveXServers: IDisposable
    {
        /// <summary>
        /// MEF composition container
        /// </summary>
        private CompositionContainer _container;

        [ImportMany]
        private IEnumerable<Lazy<IActiveXServer, IActiveXServerAttribute>> _comServers;

        private ActiveXServerContext _context;

        public ActiveXServers(ActiveXServerContext context)
        {
            try
            {
                _context = context;

                string extensionLocation = ConfigurationManager.AppSettings["ActiveXServerPlugins:Location"];
                //// An aggregate catalog that combines multiple catalogs
                var catalog = new AggregateCatalog();
                //catalog.Catalogs.Add(new AssemblyCatalog(Assembly.GetExecutingAssembly()));
                catalog.Catalogs.Add(new DirectoryCatalog(extensionLocation));
                //// Create the CompositionContainer with the parts in the catalog
                this._container = new CompositionContainer(catalog, true);
                //// Fill the imports of this object
                try
                {
                    this._container.ComposeParts(this);
                }
                catch (CompositionException)
                {
                    throw;
                }
            }
            catch (ConfigurationErrorsException)
            {
                throw;
            }
        }

        public List<ActiveXServer> GetServers()
        {
            List<ActiveXServer> servers = new List<ActiveXServer>();

            _comServers.TryForEach(plugin =>
            {
                servers.Add(new ActiveXServer()
                {                    
                    CoClassType = plugin.Metadata.CoClassType,
                    ClassID = new Guid(plugin.Metadata.CoClassType.GetCustomAttribute<GuidAttribute>().Value),
                    PrimaryInterface =plugin.Metadata.PrimaryInterface
                });
            });
            
            return servers;
        }

        public IActiveXServer GetServerInstance(ActiveXServer server)
        {
            return _comServers.Where(plugin => plugin.Metadata.CoClassType == server.CoClassType).First().Value;
        }

        [Export("LockActiveXServer", typeof(Func<int>))]
        protected int LockActiveXServer()
        {
            return _context.Lock();
        }

        [Export("UnLockActiveXServer", typeof(Func<int>))]
        protected int UnLockActiveXServer()
        {
            return _context.Unlock();
        }

        public void Dispose()
        {
            this._container.Dispose();
            GC.SuppressFinalize(this);
        }
    };
};
