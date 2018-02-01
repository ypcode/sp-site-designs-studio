import * as React from 'react';
import {
	DocumentCard,
	DocumentCardPreview,
	DocumentCardTitle,
	DocumentCardActions,
	DocumentCardType,
	ImageFit,
	IDocumentCardPreviewProps,
	Panel,
	PanelType,
	PrimaryButton,
	DefaultButton,
	MessageBar,
	MessageBarType,
	IconButton,
	ActionButton
} from 'office-ui-fabric-react';
import styles from '../SiteDesignsStudio.module.scss';
import { ISiteDesignsStudioProps, IServiceConsumerComponentProps } from '../ISiteDesignsStudioProps';
import { escape, assign } from '@microsoft/sp-lodash-subset';

import GenericObjectEditor from '../genericObjectEditor/GenericObjectEditor';
import { ISiteScriptContent, ISiteScript, SiteScriptEntitySchema } from '../../models/ISiteScript';
import { ISiteDesignsService, SiteDesignsServiceKey } from '../../services/siteDesigns/SiteDesignsService';

export interface ISiteScriptsManagerProps extends IServiceConsumerComponentProps {
	onScriptContentEdit: (siteScript: ISiteScript) => void;
}
export interface ISiteScriptsManagerState {
	siteScripts: ISiteScript[];
	isEditing: boolean;
	currentScriptEdition: ISiteScript;
	isLoading: boolean;
	hasError: boolean;
	userMessage: string;
}

export default class SiteScriptsManager extends React.Component<ISiteScriptsManagerProps, ISiteScriptsManagerState> {
	private siteDesignsService: ISiteDesignsService;

	constructor(props: ISiteScriptsManagerProps) {
		super(props);

		this.props.serviceScope.whenFinished(() => {
			this.siteDesignsService = this.props.serviceScope.consume<ISiteDesignsService>(SiteDesignsServiceKey);
		});

		this.state = {
			siteScripts: [],
			isEditing: false,
			currentScriptEdition: null,
			isLoading: false,
			hasError: false,
			userMessage: null
		};
	}

	public componentWillMount() {
		this._loadSiteScripts();
	}

	private _loadSiteScripts(setLoading: boolean = false): Promise<any> {
		if (setLoading) {
			this.setState({
				isLoading: true
			});
		}
		return this.siteDesignsService
			.getSiteScripts()
			.then((siteScripts) => {
				this.setState({
					siteScripts: siteScripts,
					isLoading: false
				});
			})
			.catch((error) => {
				this.setState({
					siteScripts: null,
					isLoading: false
				});
			});
	}

	public render(): React.ReactElement<ISiteDesignsStudioProps> {
		let { siteScripts, isEditing, currentScriptEdition, hasError, userMessage } = this.state;
		return (
			<div className={styles.siteDesignsManager}>
				{isEditing && this._renderSiteScriptPropertiesEditor(currentScriptEdition)}
				{userMessage && (
					<div className="ms-Grid-row">
						<div className="ms-Grid-col ms-sm12">
							<MessageBar
								messageBarType={hasError ? MessageBarType.error : MessageBarType.success}
								isMultiline={false}
								onDismiss={this._clearError.bind(this)}
							>
								{userMessage}
							</MessageBar>
						</div>
					</div>
				)}
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12 ms-lg2 ms-lgOffset10">
						<ActionButton iconProps={{ iconName: 'Add' }} onClick={() => this._addNewSiteScript()}>
							New
						</ActionButton>
					</div>
				</div>
				<div className="ms-Grid-row">
					{siteScripts.map((sd, ndx) => (
						<div className="ms-Grid-col ms-sm12" key={'SD_' + ndx}>
							<div className={styles.siteDesignItem}>{this._renderSiteScriptItem(sd)}</div>
						</div>
					))}
				</div>
			</div>
		);
	}

	private _addNewSiteScript() {
		this._editSiteScript({
			Id: '',
			Title: 'New Site Script',
			Content: {
				actions: [],
				bindata: {},
				version: 1
			},
			Description: '',
			Version: 1
		});
	}

	private _editSiteScript(siteScript: ISiteScript) {
		console.log('Site script edited= ', siteScript);
		this.setState({
			isEditing: true,
			currentScriptEdition: siteScript
		});
	}

	private _deleteConfirm(siteScript: ISiteScript) {
		if (confirm(`Are you sure you want to delete this Site Script '${siteScript.Title}'?`)) {
			this._deleteScript(siteScript);
		}
	}

	private _clearError() {
		let { hasError } = this.state;
		if (hasError) {
			this.setState({ hasError: false });
		}
	}

	private _renderSiteScriptPropertiesEditor(siteScript: ISiteScript) {
		let editingSiteScript = assign({}, siteScript);

		const onObjectChanged = (o) => {
			assign(editingSiteScript, o);
		};

		return (
			<Panel isOpen={true} type={PanelType.smallFixedFar} onDismiss={() => this._cancelScriptEdition()}>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12">
						<GenericObjectEditor
							readOnlyProperties={[ 'Id' ]}
							object={editingSiteScript}
							onObjectChanged={onObjectChanged.bind(this)}
							schema={SiteScriptEntitySchema}
						/>
					</div>
				</div>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm6 ms-lg4 ms-lgOffset4">
						<PrimaryButton text="Save" onClick={() => this._saveScript(editingSiteScript)} />
					</div>
					<div className="ms-Grid-col ms-sm6 ms-lg4">
						<DefaultButton text="Cancel" onClick={() => this._cancelScriptEdition()} />
					</div>
				</div>
			</Panel>
		);
	}

	private _renderSiteScriptItem(siteScript: ISiteScript) {
		return (
			<div className={styles.siteScriptItem}>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12 ms-lg8">
						<h2 className={styles.siteScriptItemTitle}>{siteScript.Title}</h2>
					</div>
					<div className="ms-Grid-col ms-sm3 ms-lg1" />
					<div className="ms-Grid-col ms-sm3 ms-lg1">
						<IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editSiteScript(siteScript)} />
					</div>
					<div className="ms-Grid-col ms-sm3 ms-lg1">
						<IconButton
							iconProps={{ iconName: 'PageEdit' }}
							onClick={() => this.props.onScriptContentEdit(siteScript)}
						/>
					</div>
					<div className="ms-Grid-col ms-sm3 ms-lg1">
						<IconButton
							iconProps={{ iconName: 'Delete' }}
							onClick={() => this._deleteConfirm(siteScript)}
						/>
					</div>
				</div>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12">
						<h4>
							Version <strong>{siteScript.Version}</strong>
						</h4>
						<p>{siteScript.Description}</p>
					</div>
				</div>
			</div>
		);
	}

	private _saveScript(siteScript: ISiteScript) {
		// If the site script is new (has no set Id)
		if (!siteScript.Id) {
      // Redirect to edit content
			this.props.onScriptContentEdit(siteScript);
		} else {
			this.setState({ isLoading: true });
			this.siteDesignsService
				.saveSiteScript(siteScript)
				.then((_) => {
					this.setState({
						isEditing: false,
						currentScriptEdition: null,
						userMessage: 'The site script has been properly saved'
					});
				})
				.then(() => this._loadSiteScripts(true))
				.catch((error) => {
					this.setState({
						isEditing: false,
						currentScriptEdition: null,
						hasError: true,
						userMessage: 'The site script cannot be properly saved'
					});
				});
		}
	}

	private _deleteScript(siteScript: ISiteScript) {
		this.setState({ isLoading: true });
		this.siteDesignsService
			.deleteSiteScript(siteScript)
			.then((_) => {
				this.setState({
					isEditing: false,
					currentScriptEdition: null,
					userMessage: 'The site script has been properly deleted'
				});
			})
			.then(() => this._loadSiteScripts(true))
			.catch((error) => {
				this.setState({
					isEditing: false,
					currentScriptEdition: null,
					hasError: true,
					userMessage: 'The site script cannot be deleted'
				});
			});
	}

	private _cancelScriptEdition() {
		this.setState({
			isEditing: false,
			currentScriptEdition: null
		});
	}
}
