# Embodied Carbon Modeling for MicroGreen

This is the open-source repository for the embodied carbon modeling section of the MicroGreen paper. 

The model uses the ACT carbon modeling tool as a foundation for silicon-level modeling, and incorporates additional modeling methodologies for other components described in the paper, including capacitors, resistors, inductors, PCBs, and connectors.

---
### Setup
 
1. **Clone the repository** with all submodules:
 
```bash
git clone git@github.com:S4AI-CornellTech/EmbodiedCarbonModeling.git
git checkout RPI-Silicon
```
 
2. **Install Python dependencies** using the setup script:
 
```bash
bash setup.sh
```

### Generate Embodied Carbon Modeling Results for MicroGreen

Ensure you are running commands within MicroGreen's virtual environment to satisfy all library requirements.
```bash
cd "${EMBODIED_CARBON_MODELING_DIR}"
mkdir -p results
python3 -m act.act_model -m act/boms/RPi_Pico_silicon_only.yaml --export-file results/RPi_Pico_silicon_only_output
python3 -m act.act_model -m act/boms/RPi_4B_silicon_only.yaml --export-file results/RPi_4B_silicon_only_output
python3 -m act.act_model -m act/boms/RPi_CM4_silicon_only.yaml --export-file results/RPi_CM4_silicon_only_output
python3 -m act.act_model -m act/boms/RPi_400_silicon_only.yaml --export-file results/RPi_400_silicon_only.yaml
python3 -m act.act_model -m act/boms/RPi_POE_Hat_silicon_only.yaml --export-file results/RPi_POE_Hat_silicon_only_output
python3 -m act.act_model -m act/boms/RPi_Zero_2WH_silicon_only.yaml --export-file results/RPi_Zero_2WH_silicon_only_output
```

For the full list of command-line arguments, run `python -m act.act_model --help`.

---

## Bill of Materials Specification

To use our model to evaluate the BoM of other MCUs and electronic devices, you can either write your own bill of materials or start from one of the existing examples in the `boms` directory.

Once your bill of materials is ready, run it with ACT using:
```bash
python -m act.act_model -m <your_bom.yaml>
```

For example, `python -m act.act_model -m act/boms/dellr740.yaml` runs one of the stock Dell R740 models.

---

## Codebase Structure

The top-level entry point is `act_model.py`, which orchestrates calculations across the underlying embodied architectural carbon models for logic, memory, storage, peripheral components, and more.

ACT currently supports the following models:

- `logic_model.py`: Application processor and digital logic embodied carbon model
- `dram_model.py`: DRAM embodied carbon capacity-based model
- `ssd_model.py`: SSD embodied carbon capacity-based model
- `hdd_model.py`: HDD embodied carbon capacity-based model
- `storage_model.py`: Storage embodied carbon model
- `op_model.py`: Operational emissions model
- `capacitor_model.py`: Capacitor embodied carbon model
- `resistor_model.py`: Resistor embodied carbon model
- `inductor_model.py`: Inductor embodied carbon model
- `connector_model.py`: Connector embodied carbon model
- `diode_model.py`: Diode embodied carbon model
- `switch_model.py`: Switch embodied carbon model
- `other_model.py`: Other miscellaneous component embodied carbon model (e.g. passive filters)
- `materials_model.py`: Frame and enclosure materials embodied carbon model
- `pcb_model.py`: Printed circuit board area-based embodied carbon model
- `battery_model.py`: Battery capacity-based embodied carbon model
- `active_model.py`: Active component embodied carbon model

Data for the architectural carbon model is drawn from sustainability literature and industry sources. Additional information can be found in the MicroGreen paper.

---

## Link to the ACT
To read the paper please visit this [link](https://dl.acm.org/doi/abs/10.1145/3470496.3527408)


### Citation
If you use `ACT`, please cite us:                                                                 

```
@inproceedings{GuptaACT2022,
author = {Gupta, Udit and Elgamal, Mariam and Hills, Gage and Wei, Gu-Yeon and Lee, Hsien-Hsin S. and Brooks, David and Wu, Carole-Jean},
title = {ACT: Designing Sustainable Computer Systems with an Architectural Carbon Modeling Tool},
year = {2022},
isbn = {9781450386104},
publisher = {Association for Computing Machinery},
address = {New York, NY, USA},
url = {https://doi.org/10.1145/3470496.3527408},
doi = {10.1145/3470496.3527408},
abstract = {Given the performance and efficiency optimizations realized by the computer systems and architecture community over the last decades, the dominating source of computing's carbon footprint is shifting from operational emissions to embodied emissions. These embodied emissions owe to hardware manufacturing and infrastructure-related activities. Despite the rising embodied emissions, there is a distinct lack of architectural modeling tools to quantify and optimize the end-to-end carbon footprint of computing. This work proposes ACT, an architectural carbon footprint modeling framework, to enable carbon characterization and sustainability-driven early design space exploration. Using ACT we demonstrate optimizing hardware for carbon yields distinct solutions compared to optimizing for performance and efficiency. We construct use cases, based on the three tenets of sustainable design---Reduce, Reuse, Recycle---to highlight future methods that enable strong performance and efficiency scaling in an environmentally sustainable manner.},
booktitle = {Proceedings of the 49th Annual International Symposium on Computer Architecture},
pages = {784–799},
numpages = {16},
keywords = {energy, sustainable computing, computer architecture, mobile, manufacturing},
location = {New York, New York},
series = {ISCA '22}
}


```

### License
ACT is MIT licensed, as found in the LICENSE file.
