# Copyright (c) Meta Platforms, Inc. and affiliates.

# This source code is licensed under the MIT license found in the
# LICENSE file in the root directory of this source tree.

import datetime
import logging
import os
import sys
import tempfile

import pint
import yaml
from .core.units import *

from .core.arg_parser import get_clean_args, get_parser
from .core.common import *

from .core.capacitor_model import CapacitorModel, DEFAULT_CP_CONFIG
from .core.carbon import Carbon, SourceType

from .core.dram_model import DEFAULT_DRAM_CONFIG, DRAMModel
from .core.hdd_model import DEFAULT_HDD_CONFIG, HDDModel

from .core.logger import log, setup_logger
from .core.logic_model import LogicModel
from .core.materials_model import DEFAULT_MATERIALS_CONFIG, MaterialsModel
from .core.op_model import OpModel
from .core.ssd_model import SSDModel
from .core.bom import *
from .core.battery_model import BatteryModel
from .core.pcb_model import DEFAULT_PCB_MODEL_FILE, PCBModel
from .core.connector_model import ConnectorModel, DEFAULT_CONNECTOR_MODEL_FILE
from .core.diode_model import DiodeModel, DEFAULT_DIODE_MODEL_FILE
from .core.switch_model import SwitchModel, DEFAULT_SWITCH_MODEL_FILE
from .core.resistor_model import ResistorModel, DEFAULT_RESISTOR_MODEL_FILE
from .core.inductor_model import InductorModel, DEFAULT_INDUCTOR_MODEL_FILE
from .core.other_model import OtherModel, DEFAULT_OTHER_MODEL_FILE
from .core.active_model import ActiveModel, DEFAULT_ACTIVE_MODEL_FILE
from .core.utils import DEFAULT_LOCATION_CONFIG, DEFAULT_SOURCE_CONFIG


class ACTModel:
    def __init__(
        self,
        out_dir: str = None,
        weight_unit=kg,
        cap_config=DEFAULT_CP_CONFIG,
        dram_config=DEFAULT_DRAM_CONFIG,
        hdd_config=DEFAULT_HDD_CONFIG,
        materials_config=DEFAULT_MATERIALS_CONFIG,
        pcb_config=DEFAULT_PCB_MODEL_FILE,
        connector_config=DEFAULT_CONNECTOR_MODEL_FILE,
        diode_config=DEFAULT_DIODE_MODEL_FILE,
        resistor_config=DEFAULT_RESISTOR_MODEL_FILE,
        switch_config=DEFAULT_SWITCH_MODEL_FILE,
        inductor_config=DEFAULT_INDUCTOR_MODEL_FILE,
        other_config=DEFAULT_OTHER_MODEL_FILE,
        active_config=DEFAULT_ACTIVE_MODEL_FILE,
        loc_ci_config=DEFAULT_LOCATION_CONFIG,
        src_ci_config=DEFAULT_SOURCE_CONFIG,
    ):
        """ACT Model object

        Args:
            out_dir: Output directory for results
            weight_unit: The unit of weight to normalize results to for reporting
            cap_config: Capacitor model configuration file
            dram_config: DRAM model configuration file
            hdd_config: HDD model configuration file
            pcb_config: PCB model configuration file
            connector_config: Connector model configuration file
            diode_config: Diode model configuration file
            switch_config: Switch model configuration file
            resistor_config: Resistor model configuration file
            other_config: Other model configuration file
            active_config: Active semiconductor model configuration file
            materials_config: Material model configuration file
            loc_ci_config: Location carbon intensity configuration file
            src_ci_config: Energy source carbon intensity configuration file

        """

        self.weight_unit = weight_unit

        if out_dir is None:
            self.out_dir = tempfile.TemporaryDirectory(prefix="act_out_").name
        else:
            self.out_dir = out_dir

        if not os.path.exists(self.out_dir):
            os.makedirs(self.out_dir, exist_ok=True)

        # load the models for each type of device
        self.logic_model = LogicModel()
        self.dram_model = DRAMModel(model_file=dram_config)
        self.ssd_model = SSDModel()
        self.hdd_model = HDDModel(model_files=hdd_config)
        self.op_model = OpModel(
            loc_ci_config=loc_ci_config, src_ci_config=src_ci_config
        )
        self.cap_model = CapacitorModel(model_file=cap_config)
        self.materials_model = MaterialsModel(model_file=materials_config)
        self.pcb_model = PCBModel(model_file=pcb_config)
        self.connector_model = ConnectorModel(model_file=connector_config) 
        self.diode_model = DiodeModel(model_file=diode_config)
        self.resistor_model = ResistorModel(model_file=resistor_config)
        self.switch_model = SwitchModel(model_file=switch_config)
        self.inductor_model = InductorModel(model_file=inductor_config)
        self.other_model = OtherModel(model_file=other_config)
        self.active_model = ActiveModel(model_file=active_config)
        self.battery_model = BatteryModel()

        # save the last settings
        self.last_op_power = None
        self.last_op_ci = None
        self.last_duty_cycle = None
        self.last_hw_lifetime = None
        self.last_bom = None

        # results attributes
        self.silicon_results = dict()
        self.passives_results = dict()
        self.materials_results = dict()

    def get_carbon(
        self,
        bom: dict,
        op_power: pint.Quantity,
        op_ci=EnergyLocation.USA,
        duty_cycle: float = 1.0,
        hw_lifetime=2 * year,
        export_file=None,
    ):
        """Calculate the aggregate carbon cost for this configuration

        Args:
            bom: Bill of materials data structure specifying the component lists and parameters
            op_power: Operating power of the device
            op_ci: Operational carbon intensity setting
            duty_cycle: Device utilization rate between 0 and 1
            hw_lifetime: Expected hardware life cycle
            export_file: Output file for results
        """

        self.last_op_power = op_power
        self.last_op_ci = op_ci
        self.last_duty_cycle = duty_cycle
        self.last_hw_lifetime = hw_lifetime
        self.last_bom = bom

        self.silicon_results = self.silicon_analysis(bom.silicon)
        self.passives_results = self.passives_analysis(bom.passives)
        self.materials_results = self.materials_analysis(bom.materials)

        op_carbon = self.op_model.get_carbon(
            lifetime=hw_lifetime, duty_cycle=duty_cycle, op_power=op_power, op_ci=op_ci
        )

        total_carbon = (
            sum(
                [
                    *self.silicon_results.values(),
                    *self.passives_results.values(),
                    *self.materials_results.values(),
                ]
            )
            + op_carbon
        )

        # export the result to report for auditing
        if export_file is None:
            # Use the input BOM filename if available
            if hasattr(bom, 'file') and bom.file:
                input_filename = os.path.splitext(os.path.basename(bom.file))[0]
                export_file = f"{self.out_dir}/act_report_{input_filename}.yaml"
            else:
                export_file = f"{self.out_dir}/act_report.yaml"
        self.export_results(export_file, total_carbon)

        return total_carbon

    def silicon_analysis(self, silicon):
        # for each device, run the carbon modeling analysis
        silicon_results = dict()

        # for each silicon item in the list, query the manufacturing cost
        for sname, silicon_data in silicon.items():
            mtype = silicon_data.model
            fab_yield = silicon_data.fab_yield
            n_ics = silicon_data.n_ics
            gpa = silicon_data.gpa
            fab_ci = silicon_data.fab_ci

            # calculate the carbon emissions for silicon devices
            if mtype is ModelType.LOGIC:
                si_carbon = self.logic_model.get_carbon(
                    logic_process=silicon_data.process,
                    area=silicon_data.area,
                    fab_yield=fab_yield,
                    n_ics=n_ics,
                    gpa=gpa,
                    fab_ci=fab_ci,
                )
            elif mtype is ModelType.DRAM:
                si_carbon = self.dram_model.get_carbon(
                    capacity=silicon_data.capacity,
                    process=silicon_data.process,
                    fab_yield=fab_yield,
                    n_ics=n_ics,
                )
            elif mtype is ModelType.FLASH:
                si_carbon = self.ssd_model.get_carbon(
                    capacity=silicon_data.capacity,
                    process=silicon_data.process,
                    fab_yield=fab_yield,
                    n_ics=n_ics,
                )
            elif mtype is ModelType.HDD:
                si_carbon = self.hdd_model.get_carbon(
                    capacity=silicon_data.capacity,
                    process=silicon_data.process,
                    fab_yield=fab_yield,
                    n_ics=n_ics,
                )
            elif mtype is ModelType.MANUAL:
                si_carbon = Carbon(
                    silicon_data.carbon / fab_yield, silicon_data.ctype
                ) + Carbon(
                    n_ics * CARBON_PER_IC_PACKAGE / fab_yield, SourceType.PACKAGING
                )
            else:
                raise NotImplementedError(
                    f"Silicon model type for {mtype} not implemented. Unable to calculate cost."
                )

            # update the results with the aggregate total
            silicon_results[sname] = si_carbon

        return silicon_results

    def passives_analysis(self, passives):
        passives_results = dict()
        for pname, pspec in passives.items():
            if pspec.category is ComponentCategory.CAPACITOR:
                carbon = self.cap_model.get_carbon(
                    ci=pspec.fab_ci,
                    ctype=pspec.type,
                    weight=pspec.weight,
                    n_caps=pspec.quantity,
                )
            elif pspec.category is ComponentCategory.CONNECTOR:  
                carbon = self.connector_model.get_carbon(
                    weight=pspec.weight,
                    connector_type=pspec.type,
                    n_connectors=pspec.quantity,
                )
            elif pspec.category is ComponentCategory.DIODE:
                carbon = self.diode_model.get_carbon(
                    weight=pspec.weight,
                    diode_type=pspec.type,
                    n_diodes=pspec.quantity,
                )
            elif pspec.category is ComponentCategory.SWITCH:
                carbon = self.switch_model.get_carbon(
                    weight=pspec.weight,
                    switch_type=pspec.type,
                    n_switches=pspec.quantity,
                )
            elif pspec.category is ComponentCategory.INDUCTOR:
                carbon = self.inductor_model.get_carbon(
                    n_inductors=pspec.quantity,
                    inductor_type=pspec.type,
                    weight=pspec.weight if pspec.weight.magnitude > 0 else None,
                )
            elif pspec.category is ComponentCategory.RESISTOR:
                carbon = self.resistor_model.get_carbon(
                    n_resistors=pspec.quantity,
                    resistor_type=pspec.type,
                )
            elif pspec.category is ComponentCategory.OTHER:
                carbon = self.other_model.get_carbon(
                    weight=pspec.weight,
                    component_type=pspec.type,
                    n_components=pspec.quantity,
                )
            elif pspec.category is ComponentCategory.ACTIVE:
                carbon = self.active_model.get_carbon(
                    weight=pspec.weight,
                    active_type=pspec.type,
                    n_components=pspec.quantity,
                )
            else:
                raise NotImplementedError(
                    f"Carbon model for component type {pspec.category} not implemented."
                )

            # add the passive device to the results and list of devices
            passives_results[pname] = carbon
        return passives_results

    def materials_analysis(self, materials):
        materials_results = dict()
        for pname, spec in materials.items():
            if spec.category in [
                ComponentCategory.FRAME,
                ComponentCategory.ENCLOSURE,
                ComponentCategory.TIN,
                ComponentCategory.BRONZE,
                ComponentCategory.PB_FREE_SOLDER,
                ComponentCategory.ALUMINUM,
            ]:
                # If no explicit type was given (type==NA), fall back to the
                # category name as the material key (e.g. "tin" → tin emission factor).
                mat = (
                    spec.type
                    if spec.type.value != "na"
                    else type(spec.type)(spec.category.value)
                )
                carbon = self.materials_model.get_carbon(
                    mat=mat, weight=spec.weight
                )
            elif spec.category is ComponentCategory.PCB:
                carbon = self.pcb_model.get_carbon(
                    area=spec.area, 
                    layers=spec.layers,
                    thickness=spec.thickness     # added for PCB thickness
                )
            elif spec.category is ComponentCategory.BATTERY:
                carbon = self.battery_model.get_carbon(capacity=spec.capacity)
            else:
                raise NotImplementedError(
                    f"Carbon model for component type {spec.category} not implemented."
                )

            # add the passive device to the results and list of devices
            materials_results[pname] = carbon
        return materials_results

    def export_results(self, export_file: str, total_carbon):
        now = datetime.datetime.now()
        export_data = dict(report_generated=now.strftime("%m/%d/%Y %H:%M:%S"))
        export_data.update(cl_args=" ".join(sys.argv))

        # export the settings used for the operational estimate
        query_dict = dict(
            op_power=str(self.last_op_power),
            op_ci=self.last_op_ci.value,
            duty_cycle=str(self.last_duty_cycle),
            hw_lifetime=str(self.last_hw_lifetime),
        )
        export_data.update(query_settings=query_dict)

        # log the total carbon
        export_data.update(total_carbon=str(total_carbon.total().to(self.weight_unit)))

        # generate the result report by category
        result_by_cat_dict = dict()
        for src in SourceType:
            result_by_cat_dict[src.name] = str(
                total_carbon.partial(src).to(self.weight_unit)
            )
        export_data.update(result_by_category=result_by_cat_dict)

        # generate the results by component
        result_by_dev_dict = dict()

        silicon_results = dict()
        for dev, carbon in self.silicon_results.items():
            dev_dict = {
                ctype.name: str(amt.to(self.weight_unit))
                for ctype, amt in carbon.carbon_by_type.items()
            }
            silicon_results[dev] = dev_dict
        result_by_dev_dict.update(silicon_results=silicon_results)

        materials_results = dict()
        for dev, carbon in self.materials_results.items():
            dev_dict = {
                ctype.name: str(amt.to(self.weight_unit))
                for ctype, amt in carbon.carbon_by_type.items()
            }
            materials_results[dev] = dev_dict
        result_by_dev_dict.update(materials_results=materials_results)

        passives_results = dict()
        for dev, carbon in self.passives_results.items():
            dev_dict = {
                ctype.name: str(amt.to(self.weight_unit))
                for ctype, amt in carbon.carbon_by_type.items()
            }
            passives_results[dev] = dev_dict
        result_by_dev_dict.update(passives_results=passives_results)

        export_data.update(result_by_device=result_by_dev_dict)

        def _float_representer(dumper, value):
            return dumper.represent_scalar("tag:yaml.org,2002:float", f"{value:.6g}")

        yaml.add_representer(float, _float_representer)
        with open(export_file, "w") as handle:
            yaml.dump(export_data, handle)
        log.info(f"ACT results exported to: {export_file}")


def main():
    # parse arguments and sanitize them
    parser = get_parser()
    args = parser.parse_args()

    # setup logging and telemetry
    loglevel = getattr(logging, args.loglevel.upper())
    setup_logger(loglevel=loglevel)

    log.info("ACT called with: " + " ".join(sys.argv))

    model_args, query_args = get_clean_args(args)

    # initialize the model
    model = ACTModel(**model_args)

    # if a bill of materials file is specified, use that instead of the cl arg values
    if args.materials is not None:
        with open(args.materials) as handle:
            bom = BOM(
                **yaml.load(handle, Loader=yaml.FullLoader),
                file=args.materials,
                material_type=model.materials_model.MaterialType,
            )
            query_args.update(bom=bom)

    # query the model for the carbon estimate
    carbon = model.get_carbon(**query_args)
    log.info(f"Total carbon for this system configuration: {carbon.total()}")

    log.info("ACT done executing...")

    return model


if __name__ == "__main__":
    main()
