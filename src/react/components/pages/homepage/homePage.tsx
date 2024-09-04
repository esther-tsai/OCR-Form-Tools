// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { SyntheticEvent } from "react";
import { connect } from "react-redux";
import { RouteComponentProps } from "react-router-dom";
import { bindActionCreators } from "redux";
import { FontIcon } from "@fluentui/react";
import { strings, interpolate } from "../../../../common/strings";
import { getPrimaryRedTheme } from "../../../../common/themes";
import IProjectActions, * as projectActions from "../../../../redux/actions/projectActions";
import IApplicationActions, * as applicationActions from "../../../../redux/actions/applicationActions";
import IAppTitleActions, * as appTitleActions from "../../../../redux/actions/appTitleActions";
import { CloudFilePicker } from "../../common/cloudFilePicker/cloudFilePicker";
import CondensedList from "../../common/condensedList/condensedList";
import Confirm from "../../common/confirm/confirm";
import "./homePage.scss";
import RecentProjectItem from "./recentProjectItem";
import { constants } from "../../../../common/constants";
import {
    IApplicationState, IConnection, IProject,
    ErrorCode, AppError, IAppSettings,
} from "../../../../models/applicationState";
import { StorageProviderFactory } from "../../../../providers/storage/storageProviderFactory";
import { decryptProject } from "../../../../common/utils";
import { toast } from "react-toastify";

export interface IHomePageProps extends RouteComponentProps, React.Props<HomePage> {
    recentProjects: IProject[];
    connections: IConnection[];
    actions: IProjectActions;
    applicationActions: IApplicationActions;
    appSettings: IAppSettings;
    project: IProject;
    appTitleActions: IAppTitleActions;
}

export interface IHomePageState {
    cloudPickerOpen: boolean;
}

function mapStateToProps(state: IApplicationState) {
    return {
        recentProjects: state.recentProjects,
        connections: state.connections,
        appSettings: state.appSettings,
        project: state.currentProject,
    };
}

function mapDispatchToProps(dispatch) {
    return {
        actions: bindActionCreators(projectActions, dispatch),
        applicationActions: bindActionCreators(applicationActions, dispatch),
        appTitleActions: bindActionCreators(appTitleActions, dispatch),
    };
}

@connect(mapStateToProps, mapDispatchToProps)
export default class HomePage extends React.Component<IHomePageProps, IHomePageState> {

    public state: IHomePageState = {
        cloudPickerOpen: false,
    };

    private newProjectRef = React.createRef<HTMLAnchorElement>();
    private deleteConfirmRef = React.createRef<Confirm>();
    private cloudFilePickerRef = React.createRef<CloudFilePicker>();

    public async componentDidMount() {
        this.props.appTitleActions.setTitle("FoTT for v2.0 has been deprecated");
        document.title = strings.homePage.title + " - " + strings.appName;
    }

    public async componentDidUpdate() {
    }

    public render() {
        return (
            <div className="app-homepage" id="pageHome">
                <div className="app-homepage-main">
                    <div className="app-banner">
                        <span className="highlight-white">Please note that this FoTT site has been deprecated <b>since Nov. 1, 2024</b> while API support for Form Recognizer v2.0 still continues until Sep. 15, 2026.</span>
                        <br /><br />
                        <span>
                            <span>Please use </span>
                            <a href="https://aka.ms/DIStudio" target="_blank" rel="noopener noreferrer">Document Intelligence Studio</a>
                            <span> for a better experience and model quality, and to keep up with the latest features. Document Intelligence Studio supports training models with any v2.1 labeled data. Refer to the </span>
                            <a href="https://aka.ms/FRMigrateGuide" target="_blank" rel="noopener noreferrer">API migration guide</a>
                            <span> to learn more about the new API to better support the long-term product roadmap and get started with the latest GA </span>
                            <a href="https://aka.ms/FRRestApiRefLatestGA" target="_blank" rel="noopener noreferrer">REST API and SDK QuickStarts</a>.
                        </span>
                        <br />
                        <span>
                            <span>To continue using v2.0 labeled data, please build the tool from </span>
                            <a href="https://aka.ms/FoTTv20GitHub">OCR-Form-Tools</a>
                            <span> or host the </span>
                            <a href="https://aka.ms/FoTTv20Alternative">docker image</a>.
                        </span>
                    </div>
                </div>
            </div>
        );
    }

    private createNewProject = (e: SyntheticEvent) => {
        this.props.actions.closeProject();
        this.props.history.push("/projects/create");

        e.preventDefault();
    }

    private handleOpenCloudProjectClick = () => {
        this.cloudFilePickerRef.current.open();
    }

    private loadSelectedProject = async (project: IProject) => {
        await this.props.actions.loadProject(project);
        this.props.history.push(`/projects/${project.id}/edit`);
    }

    private freshLoadSelectedProject = async (project: IProject) => {
        // Lookup security token used to decrypt project settings
        const projectToken = this.props.appSettings.securityTokens
            .find((securityToken) => securityToken.name === project.securityToken);

        if (!projectToken) {
            throw new AppError(ErrorCode.SecurityTokenNotFound, "Security Token Not Found");
        }

        // Load project from storage provider to keep the project in latest state
        const decryptedProject = await decryptProject(project, projectToken);
        const storageProvider = StorageProviderFactory.createFromConnection(decryptedProject.sourceConnection);
        try {
            let projectStr: string;
            try {
                projectStr = await storageProvider.readText(
                    `${decryptedProject.name}${constants.projectFileExtension}`);
            } catch (err) {
                if (err instanceof AppError && err.errorCode === ErrorCode.BlobContainerIONotFound) {
                    // try old file extension
                    projectStr = await storageProvider.readText(
                        `${decryptedProject.name}${constants.projectFileExtensionOld}`);
                } else {
                    throw err;
                }
            }
            const selectedProject = { ...JSON.parse(projectStr), sourceConnection: project.sourceConnection };
            await this.loadSelectedProject(selectedProject);
        } catch (err) {
            if (err instanceof AppError && err.errorCode === ErrorCode.BlobContainerIONotFound) {
                const reason = interpolate(strings.errors.projectNotFound.message, { file: `${project.name}${constants.projectFileExtension}`, container: project.sourceConnection.name });
                toast.error(reason, { autoClose: false });
                return;
            }
            throw err;
        }
    }

    private deleteProject = async (project: IProject) => {
        await this.props.actions.deleteProject(project);
    }
}
