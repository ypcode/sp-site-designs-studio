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
	CommandBar,
	MessageBar,
	MessageBarType,
	Spinner,
	SpinnerSize,
	DialogFooter
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
	scriptContentJson: string;
	isValidContent: boolean;
	isInvalidSchema: boolean;
	isNewScript: boolean;
	editMode: EditMode;
	isLoading: boolean;
	hasError: boolean;
	userMessage: string;
	expandedActionIndices: number[];
	allSubactionsExpanded: boolean;
	isEditingProperties: boolean;
}

export interface ISiteScriptEditorProps extends IServiceConsumerComponentProps {
	script: ISiteScript;
}

export default class SiteScriptEditor extends React.Component<ISiteScriptEditorProps, ISiteScriptEditorState> {
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
			scriptContentJson: '',
			isNewScript: false,
			isValidContent: true,
			isInvalidSchema: false,
			editMode: EditMode.Designer,
			isLoading: true,
			hasError: false,
			userMessage: '',
			expandedActionIndices: [],
			allSubactionsExpanded: false,
			isEditingProperties: false
		};
	}

	public componentWillMount() {
		let { script } = this.props;

		this.siteScriptSchemaService
			.getAvailableActionsAsync()
			.then(() => {
				// If existing script (The Id is known)
				if (script.Id) {
					// Load that script
					return this.siteDesignsService.getSiteScript(script.Id);
				} else {
					// If the argument is a new script
					// Initialize the content
					return this._initializeScriptContent(script);
				}
			})
			.then((loadedScript) => {
				// If the script content is not loaded => ERROR
				if (!loadedScript.Content && loadedScript.Id) {
					this.setState({
						script: null,
						isLoading: false,
						isValidContent: false,
						hasError: true,
						userMessage: 'The specified script is invalid'
					});
					return;
				}

				this.setState({
					script: loadedScript,
					isNewScript: loadedScript.Id ? false : true,
					isLoading: false,
					isInvalidSchema: false,
					scriptContentJson: JSON.stringify(loadedScript.Content, null, 2)
				});
			});
	}

	public render(): React.ReactElement<ISiteScriptEditorProps> {
		let { isLoading, isEditingProperties, script, isValidContent, isNewScript, hasError, userMessage } = this.state;

		if (isLoading) {
			return (
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm6 ms-smOffset3">
						<Spinner size={SpinnerSize.large} label="Loading..." />
					</div>
				</div>
			);
		}

		return (
			<div>
				{userMessage && (
					<MessageBar messageBarType={hasError ? MessageBarType.error : MessageBarType.success}>
						{userMessage}
					</MessageBar>
				)}
				{(isNewScript || isEditingProperties) && this._renderSiteScriptPropertiesEditor()}
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12 ms-lg10 ms-lgOffset1">
						<h2>{script.Title || 'No Title'}</h2>
					</div>
				</div>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12 ms-lg10 ms-lgOffset1">
						<CommandBar items={this._getCommands()} />
					</div>
				</div>
				{isValidContent && (
					<div className="ms-Grid-row">
						<div className={styles.designWorkspace}>{this._renderEditor()}</div>
					</div>
				)}
			</div>
		);
	}

	private _initializeScriptContent(script: ISiteScript): Promise<ISiteScript> {
		// TODO Fetch initial schema from Schema service
		script.Content = {
			actions: [],
			bindata: {},
			version: 1
		};
		return Promise.resolve(script);
	}
	private _getCommands() {
		let { script, editMode, isValidContent } = this.state;
		const saveBtn = {
			key: 'saveBtn',
			label: 'Save',
			title: 'Save',
			iconProps: { iconName: 'Save' },
			onClick: () => this._saveSiteScript(script)
		};
		const editBtn = {
			key: 'btnEdit',
			label: 'Edit',
			title: 'Edit Properties',
			iconProps: { iconName: 'Edit' },
			onClick: () => this._editProperties()
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

		let commands = [];

		if (isValidContent) {
			commands = commands.concat(saveBtn, editBtn);

			if (editMode == EditMode.Designer) {
				commands = commands.concat(expandAllBtn, collapseAllBtn);
			}
		}

		commands = commands.concat(switchModeBtn);

		return commands;
	}
	private _setExpanded(actionIndex: number, isExpanded: boolean) {
		let { expandedActionIndices } = this.state;
		let woCurrentIndex = expandedActionIndices.filter((i) => i != actionIndex);
		this.setState({
			expandedActionIndices: isExpanded ? woCurrentIndex.concat(actionIndex) : woCurrentIndex,
			allSubactionsExpanded: false
		});
	}

	private _setAllExpanded(isExpanded: boolean) {
		let { expandedActionIndices, script } = this.state;
		this.setState({
			expandedActionIndices: isExpanded ? script.Content.actions.map((item, index) => index) : [],
			allSubactionsExpanded: isExpanded
		});
	}

	private _isAnyCollapsed(): boolean {
		let { expandedActionIndices, script } = this.state;
		return expandedActionIndices.length < script.Content.actions.length;
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
		this.setState({
			isEditingProperties: true
		});
	}

	private _renderSiteScriptPropertiesEditor() {
		let { script } = this.state;
		let editingSiteScript = assign({}, script);

		const onObjectChanged = (o) => {
			assign(editingSiteScript, o);
		};

		// If the Id is not set, do not render it
		let ignoredProperties = editingSiteScript.Id ? [] : [ 'Id' ];

		return (
			<Panel isOpen={true} type={PanelType.smallFixedFar} onDismiss={() => this._cancelScriptPropertiesEdition()}>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12">
						<GenericObjectEditor
							ignoredProperties={ignoredProperties}
							readOnlyProperties={[ 'Id' ]}
							object={editingSiteScript}
							onObjectChanged={onObjectChanged.bind(this)}
							schema={SiteScriptEntitySchema}
						/>
					</div>
				</div>
				<DialogFooter>
					<PrimaryButton text="Ok" onClick={() => this._applyPropertiesEdition(editingSiteScript)} />
					<DefaultButton text="Cancel" onClick={() => this._cancelScriptPropertiesEdition()} />
				</DialogFooter>
			</Panel>
		);
	}

	private _renderEditor() {
		let { script, scriptContentJson, isInvalidSchema, allSubactionsExpanded } = this.state;
		switch (this.state.editMode) {
			case EditMode.Raw:
				return (
					<TextField
						multiline={true}
						rows={25}
						value={isInvalidSchema ? scriptContentJson : JSON.stringify(script.Content, null, 2)}
					/>
				);
			case EditMode.Designer:
			default:
				return (
					<div>
						<div className="ms-Grid-row">
							{script.Content.actions.map((action, index) => (
								<div>
									<ScriptActionEditor
										key={`ACTION_${index}`}
										isExpanded={this._isExpanded(index)}
										onExpandToggle={(isExpanded) => this._setExpanded(index, isExpanded)}
										allSubactionsExpanded={allSubactionsExpanded}
										serviceScope={this.props.serviceScope}
										action={action}
										getActionName={(s) => s.verb}
										schema={this.siteScriptSchemaService.getActionSchema(action)}
										onRemove={() => this._removeScriptAction(index)}
										onActionChanged={(a) => this._onActionUpdated(index, a)}
										canMoveDown={index < script.Content.actions.length - 1}
										onMoveDown={() => this._moveActionDown(index)}
										canMoveUp={index > 0}
										onMoveUp={() => this._moveActionUp(index)}
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

	private _moveActionUp(index: number) {
		this._swapActions(index, index - 1);
	}

	private _moveActionDown(index: number) {
		this._swapActions(index, index + 1);
	}

	private _swapActions(oldIndex: number, newIndex: number) {
		let { script } = this.state;
		if (newIndex < 0 || newIndex > script.Content.actions.length - 1) {
			return;
		}

		let newActions = [].concat(script.Content.actions);
		let newContent = assign({}, script.Content);
		newContent.actions = newActions;
		let newScript = assign({}, script);
		newScript.Content = newContent;

		let old = newActions[oldIndex];
		newActions[oldIndex] = newActions[newIndex];
		newActions[newIndex] = old;

		this.setState({ script: newScript });
	}

	private _addScriptAction(verb: string) {
		let { script } = this.state;
		let newAction: ISiteScriptAction = {
			verb: verb
		};

		let newActionsArray = [].concat(script.Content.actions, newAction);
		let newScriptContent = assign({}, script.Content);
		newScriptContent.actions = newActionsArray;
		let newScript = assign({}, script);
		newScript.Content = newScriptContent;
		this.setState({ script: newScript, expandedActionIndices: [ newActionsArray.length - 1 ] });
	}

	private _removeScriptAction(actionIndex: number) {
		let { script } = this.state;
		let newActionsArray = script.Content.actions.filter((item, index) => index != actionIndex);
		let newScriptContent = assign({}, script.Content);
		newScriptContent.actions = newActionsArray;
		let newScript = assign({}, script);
		newScript.Content = newScriptContent;
		this.setState({ script: newScript, expandedActionIndices: [] });
	}

	private _onActionUpdated(actionKey: number, action: ISiteScriptAction) {
		let { script } = this.state;
		let newScript: ISiteScript = assign({}, script);
		let newScriptContent = assign({}, script.Content);

		newScriptContent.actions = [].concat(newScriptContent.actions);

		// Replace the appropriate action
		newScriptContent.actions[actionKey] = action;

		newScript.Content = newScriptContent;
		this.setState({
			script: newScript
		});
	}
	private _validateForSave(siteScript: ISiteScript): string {
		if (!siteScript.Title) {
			return 'The Site Script has no title';
		}

		if (!siteScript.Content) {
			return 'The Site Script has no content';
		}

		// TODO Check content schema validity

		return null;
	}
	private _saveSiteScript(siteScript: ISiteScript) {
		let invalidMessage = this._validateForSave(siteScript);
		if (invalidMessage) {
			this.setState({
				hasError: true,
				userMessage: invalidMessage
			});
			return;
		}

		this.setState({ isLoading: true, isEditingProperties: false });
		this.siteDesignsService
			.saveSiteScript(siteScript)
			.then((_) => {
				this.setState({
					isEditingProperties: false,
					isNewScript: false,
					isLoading: false,
					hasError: false,
					userMessage: 'The site script have been properly saved'
				});
			})
			.catch((error) => {
				this.setState({
					isEditingProperties: false,
					hasError: true,
					isNewScript: false,
					isLoading: false,
					userMessage: 'The site script cannot be properly saved'
				});
			});
	}

	private _applyPropertiesEdition(siteScript: ISiteScript) {
		this.setState({
			script: siteScript,
			isEditingProperties: false,
			isNewScript: false
		});
	}
	private _cancelScriptPropertiesEdition() {
		this.setState({
			isEditingProperties: false,
			isNewScript: false
		});
	}
}
