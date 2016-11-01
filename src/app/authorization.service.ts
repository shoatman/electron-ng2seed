import { Injectable } from '@angular/core';
import {remote} from 'electron';
//import {AuthenticationContextConfig, AuthenticationContext, Logger, LogLevel, CONSTANTS} from './adal';
import * as authorayes from 'authorayes';
import * as Promise from 'bluebird';
import * as keytar from 'keytar';


/*
Notes:

- Need to detect when the user closes the popup rather than this code...

*/

interface CallBackFunc {
	(msg:string, token:any) : void;
}

class ElectronInteractiveAuthorizationCommand extends authorayes.InteractiveAuthorizationCommand {

	private authWindow: Electron.BrowserWindow;
	private expectedRedirectUri: string;
	private resolve: any;
	private reject: any;

	private createAuthWindow(){
		var webPreferences: Electron.WebPreferences = {};
		var options: Electron.BrowserWindowOptions = {};
		options.width = 800;
		options.height = 600;
		options.show = false;
		webPreferences.nodeIntegration = false;
		options.webPreferences = webPreferences;

		console.log(options);

		this.authWindow = new remote.BrowserWindow(options);

		var handleRedirect = function(event: Electron.Event, oldURL:string, newURL:string, isMainFrame:boolean, httpResponseCode:number, requestMethod:string, referrer:string, headers:Electron.Headers){
			this.handleCallBack(newURL);
			console.log('redirect detected');
		};

		var handleNavigation = function(event: Electron.Event, url:string, that:any){
			this.handleCallBack(url);
			console.log('navigation detected');
		};

		var handleRedirectBound = handleRedirect.bind(this);
		var handleNavigationBound = handleNavigation.bind(this);
		
		this.authWindow.webContents.on("did-get-redirect-request", handleRedirectBound);
		this.authWindow.webContents.on("will-navigate", handleNavigationBound);

	}

	execute(url:string, config:authorayes.InteractiveAuthorizationConfig): Promise<any>{
		this.expectedRedirectUri = config.expectedRedirectUri;
		var self: ElectronInteractiveAuthorizationCommand = this;
		this.createAuthWindow();
		this.authWindow.loadURL(url);
		this.authWindow.show();
		return new Promise(function(resolve, reject){
			//Save these for later
			self.resolve = resolve;
			self.reject = reject;
		});
	}

	private handleCallBack(url:string){
		if (url.indexOf(this.expectedRedirectUri) != -1) {
			
			//this.authContext.handleWindowCallback(url.substr(url.indexOf('#'))); //This doesn't work due to checkfor popup=true...
			var hash = url.substr(url.indexOf('#'));
			this.authWindow.destroy();
			this.resolve(hash);
			//var requestInfo = this.authContext.getRequestInfo(hash);
            //this.authContext.saveTokenFromHash(requestInfo);
            //this.authContext.callback("", requestInfo.parameters[CONSTANTS.ID_TOKEN]);
			
		}
	}
}


@Injectable()
export class AuthorizationService {

	private authContext: any;
	private authWindow: Electron.BrowserWindow;
	private armResourceId: string = "797f4846-ba00-4fd7-ba43-dac1f8f63013";
	private tokenBroker: authorayes.AADTokenBroker;

	public constructor() {


		var config: authorayes.AADTokenBrokerConfig = {
			tenantId: 'common', //COMMON OR YOUR TENANT ID
      		clientId: '188743b4-4d95-4971-9d7b-1f2c105c9bda', //REPLACE WITH YOUR CLIENT ID
      		redirectUri: 'http://armhelper', //REPLACE WITH YOUR REDIRECT URL
      		interactiveAuthorizationCommand: new ElectronInteractiveAuthorizationCommand(),
      		secureStorage: keytar,
      		storage:window.localStorage,
      		appName:"seed",
      		crypto:window.crypto
		};

		this.tokenBroker = new authorayes.AADTokenBroker(config);

	}

	public acquireToken(resource: string){
		var params: authorayes.TokenParameters = {
			resourceId: resource
		}
		this.tokenBroker.getToken(params).then(function(result:any){
			console.log('token came back');
			console.log(result);
		}).catch(function(err:any){
			console.log("error came back");
			console.log(err);
		})
	}

	
	

}