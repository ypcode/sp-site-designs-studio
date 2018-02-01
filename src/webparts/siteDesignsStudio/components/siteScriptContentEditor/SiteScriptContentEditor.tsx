import * as React from 'react';
import {
	Button,
	Dropdown,
	TextField,
	IconButton,
	Panel,
	PanelType,
	PrimaryButton,
	DefaultButton,
	CommandBar
} from 'office-ui-fabric-react';
import styles from '../SiteDesignsStudio.module.scss';
import { escape, assign } from '@microsoft/sp-lodash-subset';

import ScriptActionEditor from '../scriptActionEditor/ScriptActionEditor';
import { ISiteScriptContent, ISiteScriptAction, ISiteScript, SiteScriptEntitySchema } from '../../models/ISiteScript';
import Schema from '../../schema/schema';
import { IServiceConsumerComponentProps } from '../ISiteDesignsStudioProps';
import {
	ISiteScriptSchemaService,
	SiteScriptSchemaServiceKey
} from '../../services/siteScriptSchema/SiteScriptSchemaService';
import ScriptActionAdder from '../scriptActionAdder/ScriptActionAdder';
import { SiteDesignsServiceKey, ISiteDesignsService } from '../../services/siteDesigns/SiteDesignsService';
import GenericObjectEditor from '../genericObjectEditor/GenericObjectEditor';

// const ajv = require('ajv');
// ajv.addMetaSchema(require('ajv/lib/refs/json-schema-draft-06.json'));

export enum EditMode {
	Designer,
	Raw
}

export interface ISiteScriptEditorState {
	script: ISiteScript;
	scriptContent: ISiteScriptContent;
	scriptContentDirtyJson: string;
	isInvalidSchema: boolean;
	editMode: EditMode;
	isLoaded: boolean;
	isLoading: boolean;
	hasError: boolean;
	userMessage: string;
	expandedActionIndices: number[];
	isEditingProperties: boolean;
}

export interface ISiteScriptEditorProps extends IServiceConsumerComponentProps {
	script: ISiteScript;
}

export default class SiteScriptContentEditor extends React.Component<ISiteScriptEditorProps, ISiteScriptEditorState> {
	private siteScriptSchemaService: ISiteScriptSchemaService;
	private siteDesignsService: ISiteDesignsService;

	constructor(props: ISiteScriptEditorProps) {
		super(props);

		this.props.serviceScope.whenFinished(() => {
			this.siteScriptSchemaService = this.props.serviceScope.consume(SiteScriptSchemaServiceKey);
			this.siteDesignsService = this.props.serviceScope.consume(SiteDesignsServiceKey);
		});

		this.state = {
			script: null,
			scriptContent: null,
			scriptContentDirtyJson: '',
			isInvalidSchema: false,
			editMode: EditMode.Designer,
			isLoaded: false,
			isLoading: true,
			hasError: false,
			userMessage: '',
			expandedActionIndices: [],
			isEditingProperties: false
		};
	}

	public componentWillMount() {
		let { script } = this.props;
		this.siteScriptSchemaService
			.getAvailableActionsAsync()
			.then(() => this.siteDesignsService.getSiteScript(script.Id))
			.then((scriptWithContent) => {
				// Validate script with schema
				// let validSchema = ajv.validate(Schema, scriptWithContent);
				// if (!validSchema) {
				// 	console.log(ajv.errors);
				// 	this.setState({
				// 		script: scriptWithContent,
				// 		scriptContent: null,
				// 		isLoaded: true,
        //     isInvalidSchema: true,
        //     hasError: true,
        //     userMessage: 'The registered JSON content is not validated against JSON schema',
				// 		scriptContentDirtyJson: JSON.stringify(scriptWithContent.Content, null, 2)
				// 	});
				// } else {
					this.setState({
						script: scriptWithContent,
						scriptContent: scriptWithContent.Content,
						isLoaded: true,
						isInvalidSchema: false,
						scriptContentDirtyJson: JSON.stringify(scriptWithContent.Content, null, 2)
					});
				// }
			});
	}

	public render(): React.ReactElement<ISiteScriptEditorProps> {
		if (!this.state.isLoaded) {
			return null;
		}

		let { isEditingProperties, script } = this.state;

		return (
			<div>
				{isEditingProperties && this._renderSiteScriptPropertiesEditor()}
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12 ms-lg10 ms-lgOffset1">
						<h2>{script.Title}</h2>
					</div>
				</div>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12 ms-lg10 ms-lgOffset1">
						<CommandBar items={this._getCommands()} />
					</div>
				</div>
				<div className="ms-Grid-row">
					<div className={styles.designWorkspace}>{this._renderEditor()}</div>
				</div>
			</div>
		);
	}

	private _getCommands() {
		let { script, editMode } = this.state;
		const saveBtn = {
			key: 'saveBtn',
			label: 'Save',
			title: 'Save',
			iconProps: { iconName: 'Save' },
			onClick: () => this._saveSiteScript()
		};
		const editBtn = {
			key: 'btnEdit',
			label: 'Edit',
			title: 'Edit Properties',
			iconProps: { iconName: 'Edit' },
			onClick: () => this._saveSiteScript()
		};
		const expandAllBtn = {
			key: 'expandAllBtn',
			label: 'Expand All',
			title: 'Expand All',
			disabled: !this._isAnyCollapsed(),
			iconProps: { iconName: 'ExploreContent' },
			onClick: () => this._setAllExpanded(true)
		};
		const collapseAllBtn = {
			key: 'btnCollapseAll',
			label: 'Collapse All',
			title: 'Collapse All',
			iconProps: { iconName: 'CollapseContent' },
			onClick: () => this._setAllExpanded(false)
		};
		const switchModeBtn = {
			key: 'btnSwitchDesignRawMode',
			label: 'Switch Mode',
			title: 'Switch Design/Raw edition mode',
			iconProps: { iconName: this._getEditScriptButtonIcon() },
			onClick: () => this._toggleEditMode()
		};

		let commands = [ saveBtn ];

		if (script.Id) {
			commands = [].concat(editBtn);
		}

		if (editMode == EditMode.Designer) {
			commands = commands.concat(expandAllBtn, collapseAllBtn);
		}

		commands = commands.concat(switchModeBtn);

		return commands;
	}
	private _setExpanded(actionIndex: number, isExpanded: boolean) {
		let { expandedActionIndices } = this.state;
		let woCurrentIndex = expandedActionIndices.filter((i) => i != actionIndex);
		this.setState({
			expandedActionIndices: isExpanded ? woCurrentIndex.concat(actionIndex) : woCurrentIndex
		});
	}

	private _setAllExpanded(isExpanded: boolean) {
		let { expandedActionIndices, scriptContent } = this.state;
		this.setState({
			expandedActionIndices: isExpanded ? scriptContent.actions.map((item, index) => index) : []
		});
	}

	private _isAnyCollapsed(): boolean {
		let { expandedActionIndices, scriptContent } = this.state;
		return expandedActionIndices.length < scriptContent.actions.length;
	}

	private _isAnyExpanded(): boolean {
		let { expandedActionIndices } = this.state;
		return expandedActionIndices.length > 0;
	}

	private _isExpanded(index: number): boolean {
		return this.state.expandedActionIndices.indexOf(index) > -1;
	}

	private _getEditScriptButtonIcon() {
		let { editMode } = this.state;
		switch (editMode) {
			case EditMode.Designer:
				return 'FileCode';
			case EditMode.Raw:
				return 'Design';
		}
	}

	private _saveSiteScript() {
		let { script, scriptContent } = this.state;

		let scriptToSave = assign({}, script);
		scriptToSave.Content = assign({}, scriptContent);

		this.setState({ isLoading: true });
		this.siteDesignsService
			.saveSiteScript(scriptToSave)
			.then((_) => {
				this.setState({
					userMessage: 'The site script has been properly saved'
				});
			})
			.catch((error) => {
				this.setState({
					hasError: true,
					userMessage: 'The site script cannot be properly saved'
				});
			});
	}

	private _toggleEditMode() {
		let { editMode } = this.state;
		switch (editMode) {
			case EditMode.Designer:
				this.setState({ editMode: EditMode.Raw });
				break;
			case EditMode.Raw:
				this.setState({ editMode: EditMode.Designer });
				break;
		}
	}

	private _editProperties() {
		this.setState({ isEditingProperties: true });
	}

	private _renderSiteScriptPropertiesEditor() {
		let { script } = this.state;
		let editingSiteScript = assign({}, script);

		const onObjectChanged = (o) => {
			assign(editingSiteScript, o);
		};

		return (
			<Panel isOpen={true} type={PanelType.smallFixedFar} onDismiss={() => this._cancelScriptPropertiesEdition()}>
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
						<PrimaryButton text="Save" onClick={() => this._saveScriptProperties(editingSiteScript)} />
					</div>
					<div className="ms-Grid-col ms-sm6 ms-lg4">
						<DefaultButton text="Cancel" onClick={() => this._cancelScriptPropertiesEdition()} />
					</div>
				</div>
			</Panel>
		);
	}

	private _renderEditor() {
		let { scriptContent, scriptContentDirtyJson, isInvalidSchema } = this.state;
		switch (this.state.editMode) {
			case EditMode.Raw:
				return (
					<TextField
						multiline={true}
						rows={25}
            value={isInvalidSchema ? scriptContentDirtyJson : JSON.stringify(scriptContent, null, 2)}
					/>
				);
			case EditMode.Designer:
			default:
				return (
					<div>
						<div className="ms-Grid-row">
							{this.state.scriptContent.actions.map((action, index) => (
								<div>
									<ScriptActionEditor
										key={index}
										isExpanded={this._isExpanded(index)}
										onExpandToggle={(isExpanded) => this._setExpanded(index, isExpanded)}
										serviceScope={this.props.serviceScope}
										action={action}
										getActionName={(s) => s.verb}
										schema={this.siteScriptSchemaService.getActionSchema(action)}
										onRemove={() => this._removeScriptAction(index)}
										onActionChanged={(a) => this._onActionUpdated(index, a)}
									/>
								</div>
							))}
						</div>
						<div className="ms-Grid-row">
							<div>
								<ScriptActionAdder
									serviceScope={this.props.serviceScope}
									onActionAdded={(a) => this._addScriptAction(a)}
								/>
							</div>
						</div>
					</div>
				);
		}
	}

	// private _updateScriptContentFromJsonString(json: string) {
	// 	try {
	// 		let contentFromJsonString = JSON.parse(json);

	// 		// Validate script with schema
	// 		// let validSchema = ajv.validate(Schema, contentFromJsonString);
	// 		if (!validSchema) {
	// 			this.setState({
	// 				isInvalidSchema: true,
  //         scriptContentDirtyJson: json,
  //         hasError: true,
  //           userMessage: 'The JSON content written manually is not validated against JSON schema',
	// 			});
	// 		} else {
  //       this.setState({
  //         scriptContent: contentFromJsonString,
  //         isInvalidSchema: false,
  //         scriptContentDirtyJson: json
  //       });
  //     }

	// 	} catch (error) {
	// 		this.setState({
	// 			isInvalidSchema: true,
	// 			scriptContentDirtyJson: json
	// 		});
	// 	}
	// }

	private _addScriptAction(verb: string) {
		let newAction: ISiteScriptAction = {
			verb: verb
		};

		let newActionsArray = [].concat(this.state.scriptContent.actions, newAction);
		let newScript = assign({}, this.state.scriptContent);
		newScript.actions = newActionsArray;
		this.setState({ scriptContent: newScript, expandedActionIndices: [ newActionsArray.length - 1 ] });
	}

	private _removeScriptAction(actionIndex: number) {
		let newActionsArray = this.state.scriptContent.actions.filter((item, index) => index != actionIndex);
		let newScript = assign({}, this.state.scriptContent);
		newScript.actions = newActionsArray;
		this.setState({ scriptContent: newScript, expandedActionIndices: [] });
	}

	private _onActionUpdated(actionKey: number, action: ISiteScriptAction) {
		let newScript: ISiteScriptContent = assign({}, this.state.scriptContent);

		newScript.actions = [].concat(this.state.scriptContent.actions);

		// Replace the appropriate action
		newScript.actions[actionKey] = action;

		this.setState({
			scriptContent: newScript
		});
	}
	private _saveScriptProperties(siteScript: ISiteScript) {
		// If the site script is new (has no set Id)
		if (!siteScript.Id) {
			return; // Can only save script properties for existing script from here
		} else {
			this.setState({ isLoading: true, isEditingProperties: false });
			this.siteDesignsService
				.saveSiteScript(siteScript)
				.then((_) => {
					this.setState({
						isEditingProperties: false,
						userMessage: 'The site script properties have been properly saved'
					});
				})
				.catch((error) => {
					this.setState({
						isEditingProperties: false,
						hasError: true,
						userMessage: 'The site script properties cannot be properly saved'
					});
				});
		}
	}
	private _cancelScriptPropertiesEdition() {
		this.setState({
			isEditingProperties: false
		});
	}
}
