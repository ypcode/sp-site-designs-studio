import { ServiceScope } from '@microsoft/sp-core-library';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SiteDesignsServiceKey, ISiteDesignsService } from './services/siteDesigns/SiteDesignsService';
import { MockSiteDesignsService } from './services/siteDesigns/MockSiteDesignsService';

export class AppStartup {
	public static configureServices(appContext: IWebPartContext): Promise<ServiceScope> {
    console.log("Env=", Environment.type);
		switch (Environment.type) {
			case EnvironmentType.Local:
			case EnvironmentType.Test:
				return AppStartup.configureTestServices(appContext);
			default:
				return AppStartup.configureProductionServices(appContext);
		}
	}

	private static configureTestServices(appContext: IWebPartContext): Promise<ServiceScope> {
		return new Promise((resolve, reject) => {
			let rootServiceScope = appContext.host.serviceScope;
			rootServiceScope.whenFinished(() => {
				// Here create a dedicated service scope for test or local context
				const childScope: ServiceScope = rootServiceScope.startNewChild();
				// Register the services that will override default implementation
				childScope.createAndProvide(SiteDesignsServiceKey, MockSiteDesignsService);
				// Must call the finish() method to make sure the child scope is ready to be used
				childScope.finish();

				childScope.whenFinished(() => {
					// If other services must be used, it must done HERE
					resolve(childScope);
				});
			});
		});
	}

	private static configureProductionServices(appContext: IWebPartContext): Promise<ServiceScope> {
		return new Promise((resolve, reject) => {
			let rootServiceScope = appContext.host.serviceScope;
			rootServiceScope.whenFinished(() => {
        // Configure the service with the current context url
        let siteDesignsService: ISiteDesignsService = rootServiceScope.consume<ISiteDesignsService>(SiteDesignsServiceKey);
        siteDesignsService.baseUrl = appContext.pageContext.web.serverRelativeUrl;
        resolve(rootServiceScope);
      });
		});
	}
}
