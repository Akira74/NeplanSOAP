#-------------------------------------------------------------------------------
# Name:         Neplan 10 WS
# Purpose:      Python wrapper around Neplan 10 webservice
#
# Author:      kristjan.vilgo
#
# Created:     07.02.2019
# Copyright:   (c) kristjan.vilgo 2019
# Licence:     GPLv2
#-------------------------------------------------------------------------------
from __future__ import print_function

import inspect
import re

from pathlib import Path

#from hashlib import md5 # Before Neplan 10.8.2.0
from hashlib import sha1
from uuid import uuid4

from zeep import Transport, Client, settings
from zeep.wsse import UsernameToken
from zeep.plugins import HistoryPlugin

from requests import Session
from lxml import etree

from datetime import datetime

from urllib.parse import urlparse, urlunparse

import pandas

from enum import Enum

import urllib3
urllib3.disable_warnings()
settings.Settings(strict=False)

import aniso8601


pandas.set_option("display.max_rows", 10)
pandas.set_option("display.max_columns", 12)
pandas.set_option("display.width", 1500)
#pandas.set_option('precision', 1)


# --- FUNCTIONS ---

class NeplanService():

    # HELPER FUNCTIONS - START

    def __init__(self, server, username, password, debug = False):
        """Sets up the Neplan SOAP WS and retuns the service object
        use service.history to get last sent and received raw SOAP messages"""

        self.username = username
        self.server = server
        self.debug  = debug

        # Create crypto container for password
        #crypted_container = md5() # Before Neplan 10.8.2.0
        crypted_container = sha1()
        crypted_container.update(password.encode())
        crypted_password = crypted_container.hexdigest()

        # Suppress certificate validation
        session = Session()
        session.verify = False # to enable http and non certified https connections

        # Setup of transport
        transport = Transport(session=session)#, timeout=10, operation_timeout=360)

        # Add plugin for message exchange history
        self.history = HistoryPlugin() # Call this element to see last sent/recieved messages

        # Set up service
        wsdl = "{}/Services/External/NeplanService.svc?singleWsdl".format(server)
        wsse = UsernameToken(username, password=crypted_password)

        client = Client(wsdl, transport=transport, wsse=wsse, plugins=[self.history])
        self.client = client
        client.debug = debug # Only on Kristjan machine this has effect (prints out all sent and recived messages, direct modification to zeep libary)

        service = client.create_service('{http://www.neplan.ch/Web/External}BasicHttpBinding_NeplanService', '{}/Services/External/NeplanService.svc/basic'.format(server))

        self.service = service

        if debug:
            print("INFO - Service created to {}".format(server))


    def print_last_messageexchange(self):

        """Prints out last sent and recieved SOAP messages"""

        messages = {"SENT":     self.history.last_sent,
                    "RECIEVED": self.history.last_received}

        for message in messages:

            print("---{}---".format(message))
            print("### http header ###")
            print('\n' * 1)
            print(messages[message]["http_headers"])
            print('\n' * 1)
            print("### {} http envelope START ###".format(message))
            print('\n' * 1)
            print(etree.tostring(messages[message]["envelope"], pretty_print = True).decode())
            print("### {} http envelope END ###".format(message))
            print('\n' * 1)


    def print_duration(self, text, start_time):

        """Print duration between now and start time
        Input: text, start_time
        Output: duration (in seconds), end_time"""

        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()
        print(text, duration)

        return duration, end_time


    def update_url_to_current_server(self, localhost_url):

        """Updates localhost url to server url"""

        parsed_url      = urlparse(localhost_url)
        new_netloc      = urlparse(self.server).netloc
        parse_update    = parsed_url._replace(netloc=new_netloc)
        updated_url     = urlunparse(parse_update)

        return updated_url


    # HELPER FUNCTIONS - END


    def WriteMessageToLogFile(self, project, message_text, log_level_text = "Info"):
        """Writes a message to the user log file, by default the log level is Info
        WriteMessageToLogFile(project: ns2:ExternalProject, text: xsd:string, logLvl: xsd:string)"""

        self.service.WriteMessageToLogFile(project, message_text, log_level_text)

    # VARIABLES start
    class networkTypeGroup(Enum):
        Network             = 0
        Feeder              = 1
        Zone                = 2
        Area                = 3
        Subarea             = 4
        TieFlowBetweenZones = 5
        TieFlowBetweenAreas = 6
        VoltageLevel        = 7

    class phase(Enum):
        L1L2L3N     = 0
        L1N         = 1
        L2N         = 2
        L3N         = 3
        L1L2N       = 4
        L1L3N       = 5
        L2L3N       = 6
        L1L2L3N_AS  = 7

    class portNr(Enum):
        first   = 0
        second  = 1
        third   = 2
        fourth  = 3

    class units(Enum):
        HV = 0
        LV = 1

    class analysisType:

        __DistrictHeatingAnalysis_string    = """DistrictHeating, DistrictHeating2DProfile, DistrictHeatingContingencyAnalysis, DistrictHeatingTimeSimulation, DistrictHeatingTransientAnalysis"""
        DistrictHeatingAnalysis             = Enum('DistrictHeatingAnalysis', dict((item, item) for item in re.sub('\s+', '',__DistrictHeatingAnalysis_string).split(",")))

        __GasAnalysis_string                = """Gas, Gas2DProfile, GasContingencyAnalysis, GasTimeSimulation, GasTransientAnalysis, InteractiveDistributionGW, PressureProfileGW, TimeDistributionGW"""
        GasAnalysis                         = Enum('GasAnalysis', dict((item, item) for item in re.sub('\s+', '',__GasAnalysis_string).split(",")))

        __PowerSystemAnalysis_string        = """ArcFlash, CableSizing,
                                                CapacitorPlacement, CircuitBreakerPlacement, ContingencyAnalysis,
                                                DistanceProtection, DynamicAnalysis, ExpressFeeder, FaultFinding,
                                                FeederReinforcement, FlickerAnalysis, GroundingSystemAnalysis,
                                                HarmonicAnalysis, HostingCapacity, InvestmentAnalysis, LoadFlow,
                                                LoadFlowTimeSimulation, LowVoltageCalculation,
                                                MotorStartingAnalysis, NeplanDach, NetworkReduction, NTC,
                                                OptimalPowerFlow, OvercurrentProtection, PhaseSwapping,
                                                PipeTypeOptimization, PoleStrengthCalculation, PortionOfFeeders,
                                                PowerSystemAssessmentERIS, Reliability, Resupply, ShortCircuit,
                                                SmallSignalStability, SwitchingOptimization, ThermalAnalysis,
                                                VoltageRegulatorPlacement, VoltageStability"""
        PowerSystemAnalysis                 = Enum('PowerSystemAnalysis', dict((item, item) for item in re.sub('\s+', '', __PowerSystemAnalysis_string).split(",")))

        __WaterAnalysis_string              = """FireWater, InteractiveDistributionGW,
                                                PressureProfileGW, TimeDistributionGW, Water, Water2DProfile,
                                                WaterContingencyAnalysis, WaterTimeSimulation,
                                                WaterTransientAnalysis"""
        WaterAnalysis                       = Enum('WaterAnalysis', dict((item, item) for item in re.sub('\s+', '', __WaterAnalysis_string).split(",")))


    class elementType:

        __DistrictHeatingAnalysis_string    = """HeatingCentrifugalPump,
                                                HeatingCirculationPump, HeatingFitting, HeatingLoad,
                                                HeatingNode, HeatingPipe, HeatingPlant, HeatingPressureRegulator,
                                                HeatingShutOffValve, HeatingSpecialLoad, HeatingStation,
                                                HeatingValve"""
        DistrictHeatingAnalysis             = Enum('DistrictHeatingAnalysis', dict((item, item) for item in re.sub('\s+', '',__DistrictHeatingAnalysis_string).split(",")))

        __GasAnalysis_string                = """GasCentrifugalPump, GasCirculationPump,
                                                GasFitting, GasNode, GasPipe, GasPressureRegulator,
                                                GasShutOffValve, GasSpecialLoad, GasStation, GasValve,
                                                Trafo2WindingAsym"""
        GasAnalysis                         = Enum('GasAnalysis', dict((item, item) for item in re.sub('\s+', '',__GasAnalysis_string).split(",")))

        __PowerSystemAnalysis_string        = """ACCompressedAirEnergyStorage,
                                                ACDisperseGenerator, ACFlyWheel, AsynchronousMachine, Busbar,
                                                BusbarCoupler, CircuitBreaker, CircuitBreakerOnElem,
                                                CompositeLoad, CurrentTransformer, CustomerConnection,
                                                DCBattery, DCConverter, DCConverter3Pole, DCFlyWheel,
                                                DCFuelCell, DCGround, DCLine, DCLoad, DCMotor, DCNode,
                                                DCPhotoVoltaic, DCReactor, DCShunt, DCVoltageSource, DFIG,
                                                DisconnectSwitch, DisconnectSwitchOnElem, DistanceRelais,
                                                EarthingSystem, EarthSwitch, EnergyStorage, EquivalentSerieLF,
                                                EquivalentSerieSC, EquivalentShuntLF, EquivalentShuntSC,
                                                ExternalGrid, FaultIndicator, Filter, FrequencyRelais, FunctionBlock,
                                                Fuse, GenericModel, GroundElement, HarmonicCurrentSource,
                                                HarmonicVoltageSource, Inertia, Line, LineAsym, LineCoupling,
                                                LineSection, Load, LoadSwitch, LoadSwitchOnElem,
                                                MeasurementDevice, MechanicalLoad, MinMaxRelaisOnLink,
                                                MinMaxRelaisOnNode, MultiFunctionProtection, NestedBlockCCT,
                                                OvercurrentRelais, ParallelRLC, PoleSlipRelais, PowerRelais, PWM,
                                                PWM3Pole, Pylon, Reactor, Regulator, SerieEarthRLC, SerieRLC,
                                                SerieTransformator, Shunt, STATCOM, Station, SurgeArrester, SVC,
                                                SynchronousMachine, Table, TCSC, Trafo2Winding, Trafo3Winding,
                                                Trafo4Winding, TrafoRegulator, UPFC, UserDefinedPort0,
                                                UserDefinedPort1, UserDefinedPort2, UserDefinedPort3,
                                                UserDefinedPort4, VoltageRelais, VoltageTransformer"""
        PowerSystemAnalysis                 = Enum('PowerSystemAnalysis', dict((item, item) for item in re.sub('\s+', '', __PowerSystemAnalysis_string).split(",")))

        __WaterAnalysis_string              = """WaterCentrifugalPump, WaterCirculationPump,
                                                WaterFitting, WaterHydrant, WaterNode, WaterPipe,
                                                WaterReservoir, WaterShutOffValve, WaterSpecialLoad,
                                                WaterStation, WaterValve"""
        WaterAnalysis                       = Enum('WaterAnalysis', dict((item, item) for item in re.sub('\s+', '', __WaterAnalysis_string).split(",")))

    # NATIVE FUNCTIONS - START


    def GetAllFeeders(self, project):
        """Get all feeders of the project"""
        return self.service.GetAllFeeders(project)

    def GetAllSubAreas(self, project):
        """Get all subareas of the project"""
        return self.service.GetAllSubAreas(project)

    def GetAllZones(self, project):
        """Get all zones of the project"""
        return self.service.GetAllZones(project)

    def GetAllElementResults(self, project, analysisType = "LoadFlow"):
        """Gets a list of all element resultst"""
        return self.service.GetAllElementResults(project, analysisType)

    def GetAllElementsOfElementType(self, project, elementType="Line"):
        """Gets a list of all elements of the selected element type in a project"""
        # TODO add validation of element types
        return self.service.GetAllElementsOfElementType(project, elementType, {}, {})

    def GetAllElementsOfProject(self, project):
        """Returns all elements of given project in a dataframe, with columns [ID, CIM_ID, EIC_ID, NAME, TYPE]"""

        # Get all element data
        CIM_ID_dict = self.service.GetNeplanIDtoCimIDDictionary(project)
        EIC_ID_dict = self.service.GetNeplanIDtoEICodeDictionary(project)
        Name_Type_dict = self.service.GetAllElementsOfProject(project, {}, {})

        # Parse and merge all the data
        CIM_ID = pandas.DataFrame([(item.Key, item.Value) for item in CIM_ID_dict], columns=["ID", "CIM_ID"])
        EIC_ID = pandas.DataFrame([(item.Key, item.Value) for item in EIC_ID_dict], columns=["ID", "EIC_ID"])
        NAME = pandas.DataFrame([(item.Key, item.Value) for item in Name_Type_dict.elementNames.KeyValueOfstringstring], columns=["ID", "NAME"])
        TYPE = pandas.DataFrame([(item.Key, item.Value) for item in Name_Type_dict.elementTypes.KeyValueOfstringstring], columns=["ID", "TYPE"])

        return CIM_ID.merge(EIC_ID, on='ID').merge(NAME, on='ID').merge(TYPE, on='ID')

    def GetAnalysisResultFile(self, fileName):
        """Retruns analysis result file defined in: analysis_result.ResultFilename
        GetAnalysisResultFile(fileName: xsd:string) -> GetAnalysisResultFileResult: ns5:StreamBody"""

        analysis_result_string = self.service.GetAnalysisResultFile(fileName)
        #print(analysis_result_string) # Quite big xml, not reasonable to export it

        return analysis_result_string

    def GetAnaylsisLogFile(self, fileName):
        """Retruns analysis log file defined in: analysis_result.LogFilename
        GetAnaylsisLogFile(fileName: xsd:string) -> GetAnaylsisLogFileResult: ns5:StreamBody"""

        return self.service.GetAnaylsisLogFile(fileName)


    def GetCalcParameterAttributes(self, project, analysisType="LoadFlow"):
        """Returns parameters  of  the  given  analysis  type  for the given project"""
        # TODO validate the analysisType
        return api.service.GetCalcParameterAttributes(project, analysisType)

    def GetCalcParameterAttributesDescription(self, analysisType="LoadFlow"):
        """Returns parameters  of  the  given  analysis  type  for the given project"""
        # TODO validate the analysisType
        # TODO report that description is missing
        return api.service.GetCalcParameterAttributesDescription(analysisType)


    def GetProject(self, projectName= "", variantName= "",  diagramName= "", layerName= ""):

        """Gets a project based on the name
           GetProject(projectName: xsd:string, variantName: xsd:string, diagramName: xsd:string, layerName: xsd:string) -> GetProjectResult: ns2:ExternalProject"""

        if self.debug:
            print("INFO - Getting project: {}".format(projectName))

        project = self.service.GetProject(projectName, variantName,  diagramName, layerName)

        if project is None or project.ProjectID is None:

            print(locals())
            print('ERROR - Project not found')

        return project


    def GetLogFileAsList(self, print_log=False):
        """Returns whole user activity logfile as a list"""
        log_list = self.service.GetLogFileAsList()

        if print_log:
            for entry in log_list:
                print(entry)

        return log_list


    def GetLogFileAsString(self):
        """Returns whole user activity logfile as a string"""

        return self.service.GetLogFileAsString()

    def GetLogOnSessionID(self, project=""):
        """Get the session id for login to NEPLAN. Add the session id to the base url of NEPLAN"""

        return self.service.GetLogOnSessionID(project)

    def GetLogOnUrl(self):
        """returns logon url for the current user session
        GetLogOnUrl() -> GetLogOnUrlResult: xsd:string"""

        localhost_url = self.service.GetLogOnUrl()

        updated_url = self.update_url_to_current_server(localhost_url)

        return updated_url


    def GetLogOnUrlWithProject(self, project):
        """returns logon url for the current user session and project
        GetLogOnUrlWithProject(project: ns2:ExternalProject) -> GetLogOnUrlWithProjectResult: xsd:string"""

        localhost_url = self.service.GetLogOnUrlWithProject(project)

        updated_url = self.update_url_to_current_server(localhost_url)

        return updated_url


    def AnalyseVariant(self, project, analysisRefenceID= str(uuid4()), analysisModule = "LoadFlow", calcNameID = "", analysisMethode = "", conditions = "", analysisLoadOptionXML = ""):

        """ This function runs selected analyses on loaded project, by default LoadFlow -> returns: analysis_variant_result
        AnalyseVariant(project: ns2:ExternalProject, analysisRefenceID: xsd:string, analysisModule: xsd:string, calcNameID: xsd:string, analysisMethode: xsd:string, conditions: xsd:string, analysisLoadOptionXML: xsd:string) -> AnalyseVariantResult: ns2:AnalysisReturnInfo"""

        analysis_variant_result = self.service.AnalyseVariant(project, analysisRefenceID, analysisModule, calcNameID, analysisMethode, conditions, analysisLoadOptionXML)

        return analysis_variant_result

    def GetSubAreaIDByName(self, project, subAreaName):
        """Get the subarea ID of the given subarea name"""
        return self.service.GetSubAreaIDByName(project, subAreaName)

    def GetSubAreaNameByID(self, project, subAreaID):
        """Get the subarea name of the given subarea ID"""
        return self.service.GetSubAreaNameByID(project, subAreaID)

    def GetZoneIDByName(self, project, zoneName):
        """Get the Zone ID of the given Zone name"""
        return self.service.GetZoneIDByName(project, zoneName)

    def GetZoneNameByID(self, project, zoneID):
        """Get the Zone name of the given Zone ID"""
        return self.service.GetZoneNameByID(project, zoneID)






    # NATIVE FUNCTIONS - END

    # CUSTOM FUNCTIONS - START

    def run_loadflow(self, project_name, operational_state_name = ""):
        """Run basic loadflow analyses, operational state name is optional.

        Input : project_name, operational_state_name = ""
        Output: results_xml, analysis_response, project, process_log"""

        # START TIMER
        start_time = datetime.now()

        # Get project and run loadflow
        project           = self.GetProject(project_name)
        _,start_time = self.print_duration("Project Loaded -> ", start_time)

        analysis_response = self.AnalyseVariant(project, analysisModule = "LoadFlow", calcNameID = operational_state_name)
        _,start_time = self.print_duration("Load Flow finished -> ", start_time)

        # Get analysis process log
        process_log = self.GetAnaylsisLogFile(analysis_response.LogFilename)
        _,start_time = self.print_duration("Log file retrived -> ", start_time)

        # Get analysis result log
        results_xml = self.GetAnalysisResultFile(analysis_response.ResultFilename)
        _,start_time = self.print_duration("Results file retrived -> ", start_time)


        if results_xml:

            results_xml = etree.fromstring(results_xml)

        else:
            project_logon_url = self.GetLogOnUrlWithProject(project)
            print("No XML results returned, you haven't enabled 'Write XML result file' under Parameters->Storage/Messages.\n You can use this url to open the project: {}".format(project_logon_url))


        return results_xml, analysis_response, project, process_log

    def CIMExport(self, project, file_path="Export.zip",
                  ENTSOEZIP=True,
                  ExportEQ=True,
                  ExportSSH=True,
                  ExportTP=True,
                  ExportSV=True,
                  ExportMerged=False,
                  ExportDL=False,
                  ExportGL=False,
                  ExportDY=False,
                  ExportSVShortCircuit=False,
                  BoundaryPath=None,
                  BoundaryAreaName="EU",
                  DynamicLineRatingPath=None,
                  MAS="",
                  Period="1D",
                  ScenarioDateTime=datetime.utcnow(),
                  Version="001",
                  Description="Neplan Export",
                  ExportAsCGMES3=False,
                  ExportBoundary=False,
                  FileHeaderComment="OPDE Confidential",
                  IsAutomatedExport=True,
                  EqFileCIMID=False,
                  KeepEQIDConstant=True,
                  AreasToExport=[],
                  AreasToExportNames=[],
                  ListOfMASForSVExport=[],
                  BalticCGMArea=None,
                  BalticRSCExport=False,
                  ExcludeBRELL=True,
                  runPowerFlow=False,
                  operationalState=None
                  ):

        """Performs CIM export on the specified project, exports all CIM files to defined filepath, by default 'Export.zip'"""

        if BoundaryPath:
            with open(BoundaryPath, "rb") as file_object:

                print(f"Uploading boundary to Neplan {BoundaryPath}")
                BoundaryPath = self.service.ZipUpload(stream=file_object.read())
                print(BoundaryPath)


        #ns13:CimExportOptions(AreasToExport: ns4:ArrayOfguid, AreasToExportNames: ns4:ArrayOfstring, BalticCGMArea: xsd:string, BalticRSCExport: xsd:boolean, BoundaryAreaName: xsd:string, BoundaryPath: xsd:string, Description: xsd:string, DynamicLineRatingPath: xsd:string, ENTSOEZIP: xsd:boolean, EqFileCIMID: xsd:string, ExcludeBRELL: xsd:boolean, ExportAsCGMES3: xsd:boolean, ExportBoundary: xsd:boolean, ExportDL: xsd:boolean, ExportDY: xsd:boolean, ExportEQ: xsd:boolean, ExportGL: xsd:boolean, ExportMerged: xsd:boolean, ExportSSH: xsd:boolean, ExportSV: xsd:boolean, ExportSVShortCircuit: xsd:boolean, ExportTP: xsd:boolean, FileHeaderComment: xsd:string, IsAutomatedExport: xsd:boolean, KeepEQIDConstant: xsd:boolean, ListOfMASForSVExport: ns4:ArrayOfKeyValueOfstringArrayOfstringty7Ep6D1, MAS: xsd:string, Period: xsd:string, ScenarioDateTime: xsd:dateTime, Version: xsd:string)

        CIMOptions = {  'AreasToExport': AreasToExport, #ns4:ArrayOfguid
                        'AreasToExportNames': AreasToExportNames, #ns4:ArrayOfstring
                        'BalticCGMArea': BalticCGMArea, #xsd:string
                        'BalticRSCExport': BalticRSCExport, #xsd:boolean
                        'BoundaryAreaName': BoundaryAreaName, #xsd:string
                        'BoundaryPath': BoundaryPath, #xsd:string
                        'Description': Description, #xsd:string
                        'DynamicLineRatingPath': DynamicLineRatingPath, #xsd:string
                        'ENTSOEZIP': ENTSOEZIP, #xsd:boolean
                        'EqFileCIMID': EqFileCIMID, #xsd:string
                        'ExcludeBRELL': ExcludeBRELL, #xsd:boolean
                        'ExportAsCGMES3': ExportAsCGMES3, #xsd:boolean
                        'ExportBoundary': ExportBoundary, #xsd:boolean
                        'ExportDL': ExportDL, #xsd:boolean
                        'ExportDY': ExportDY, #xsd:boolean
                        'ExportEQ': ExportEQ, #xsd:boolean
                        'ExportGL': ExportGL, #xsd:boolean
                        'ExportMerged': ExportMerged, #xsd:boolean
                        'ExportSSH': ExportSSH, #xsd:boolean
                        'ExportSV': ExportSV, #xsd:boolean
                        'ExportSVShortCircuit': ExportSVShortCircuit, #xsd:boolean
                        'ExportTP': ExportTP, #xsd:boolean
                        'FileHeaderComment': FileHeaderComment, #xsd:string
                        'IsAutomatedExport': IsAutomatedExport, #xsd:boolean
                        'KeepEQIDConstant': KeepEQIDConstant, #xsd:boolean
                        'ListOfMASForSVExport': ListOfMASForSVExport, #ns4:ArrayOfKeyValueOfstringArrayOfstringty7Ep6D1
                        'MAS': MAS, #xsd:string
                        'Period': Period, #xsd:string
                        'ScenarioDateTime': ScenarioDateTime, #xsd:dateTime
                        'Version': Version, #xsd:string
}

        data = self.service.CIMExport(project, CIMOptions, operationalState=operationalState, runPowerFlow=runPowerFlow)

        with open(file_path, "wb") as file_object:
            print("INFO - exporting CIM data to {}".format(file_path))
            written_bytes = file_object.write(data)

        if written_bytes == 0:
            print("ERROR - Exported file is empty: {}".format(file_path))
            return False
        else:
            return True

    def CIMImport_from_files(self, inputFiles, projectName=None, isLocalPath=False):
        """Import CIM files to Neplan from local path"""

        # If no project name has been provided, create one automatically based on first provided filename
        if not projectName:
            projectName = f"{inputFiles[0]}_{uuid4()}"

        uploded_files = []
        for file_path_string in inputFiles:
            file_path = Path(file_path_string)
            print(f"Importing to {projectName} file {file_path_string}")

            if file_path.exists():
                with file_path.open("rb") as file_object:
                    uploded_files.append(self.service.ZipUpload(stream=file_object.read()))
            else:
                print(f"Could not find {file_path}")

        response = self.service.CIMImport(inputFiles={"string": uploded_files},
                                          isLocalPath=isLocalPath,
                                          projectName=projectName,
                                          userName=self.username)


        return response



    def CIMImport(self, projectName, inputFiles, isLocalPath=False):
        """Import CIM files to Neplan"""

        #array_of_strings = [{"string": file_path} for file_path in inputFiles]

        response = self.service.CIMImport(inputFiles={"string":inputFiles},
                                           isLocalPath=isLocalPath,
                                           projectName=projectName,
                                           userName=self.username)
        return response



if __name__ == "__main__":



    api = NeplanService(server, username, password, debug=True)

    model_name = f"Elering_EMS_BaseModel_{datetime.now():%Y%m%dT%H%M}"

    upload_response = api.CIMImport_from_files([r"C:\Users\kristjan.vilgo\Documents\GitHub\USVDM\Tools\RDF_PARSER\examples\Export.zip"], model_name)

    project = api.GetProject(upload_response.actualCreatedProjectName)

    lf_result = api.service.AnalyseVariant(project, analysisModule="LoadFlow", analysisReferenceID=str(uuid4()))

    #download_response = api.CIMExport(project=project, file_path=f"{model_name}.zip", )

    validation_response = api.service.ValidateCGMESModel(externalProject=project, tsoName="Elering")

    download_response = api.CIMExport(
        project=project,
        file_path=f"{model_name}.zip",
        Period="YR",
        MAS="http://www,elering.ee/OperationalPlanning",
        Version="001",
        ScenarioDateTime=aniso8601.parse_datetime("2023-10-18T08:30"),
        #BoundaryPath=r"C:\Users\kristjan.vilgo\Downloads\20220902T0000Z__ENTSOE_BD_001\20220902T0000Z__ENTSOE_BD_001.zip",
        BoundaryPath=r"\\elering.sise\teenused\NMM\data\ACG\BOUNDARY\20190624T0000Z__ENTSOE_BD_001.zip",
        #AreasToExportNames=["Estonia"],
        ExportBoundary=True)





    #project = api.GetProject("LoadFlow")

    #elements = api.GetAllElementsOfProject(project)
    #print(elements)
    #print("INFO - Elements count in project:")
    #print(elements.TYPE.value_counts())


    # api.service.CIMImport

    #folderpath = "D:\ACG"
    #mAS = "EESTI"
    #project = api.GetProject("ELERING AP")

    #response = api.service.ImportIEC_62325_451_2FilesAndDoCIMExport(folderpath=folderpath, variantID=project.VariantID, username=username, onlyOneEQ=0, mAS=mAS)



    #api.CIMExport(project)

    #python - mzeep < wsdl >




