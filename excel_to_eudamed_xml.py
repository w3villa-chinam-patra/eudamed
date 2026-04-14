#!/usr/bin/env python3
"""Generate EUDAMED XML from Excel and validate against the XSD schema.

This version builds the XML tree directly from Excel data and does not depend on
an XML template file.
"""

from __future__ import annotations

import argparse
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
from lxml import etree


NS = {
    "message": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Message/v1",
    "service": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Service/v1",
    "device": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/v1",
    "e": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/v1",
    "budi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/BasicUDI/v1",
    "udidi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/UDIDI/v1",
    "commondevice": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/CommonDevice/v1",
    "lngs": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Common/LanguageSpecific/v1",
    "mktinfo": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/MktInfo/MarketInfo/v1",
    "lnks": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Links/v1",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

XML_NS_MAP = {
    "message": NS["message"],
    "service": NS["service"],
    "device": NS["device"],
    "e": NS["e"],
    "budi": NS["budi"],
    "udidi": NS["udidi"],
    "commondevice": NS["commondevice"],
    "lngs": NS["lngs"],
    "mktinfo": NS["mktinfo"],
    "lnks": NS["lnks"],
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

# XSD-safe defaults used only when required fields are blank.
DEFAULT_CONVERSATION_ID = "00000000-0000-0000-0000-000000000000"
DEFAULT_CORRELATION_ID = "00000000-0000-0000-0000-000000000001"
DEFAULT_MESSAGE_ID = "00000000-0000-0000-0000-000000000002"
DEFAULT_RECIPIENT_NODE_ACTOR_CODE = "IT-MF-000000001"
DEFAULT_SENDER_NODE_ACTOR_CODE = "EUDAMED_MDR"
DEFAULT_MF_ACTOR_CODE = "IT-MF-000000001"
DEFAULT_UDI_STATUS_CODE = "ON_THE_MARKET"
DEFAULT_MDN_CODE = "A0000"
DEFAULT_REFERENCE_NUMBER = "REF-001"


RISK_CLASS_MAP = {
    "class 1": "CLASS_I",
    "class i": "CLASS_I",
    "class 2a": "CLASS_IIA",
    "class iia": "CLASS_IIA",
    "class 2b": "CLASS_IIB",
    "class iib": "CLASS_IIB",
    "class 3": "CLASS_III",
    "class iii": "CLASS_III",
}

STATUS_MAP = {
    "on the eu market": "ON_THE_MARKET",
    "on the market": "ON_THE_MARKET",
    "no longer placed on the eu market": "NO_LONGER_PLACED_ON_THE_MARKET",
    "no longer placed on the market": "NO_LONGER_PLACED_ON_THE_MARKET",
    "not intended for the eu market": "NOT_INTENDED_FOR_EU_MARKET",
}

LANGUAGE_MAP = {
    "all languages": "ANY",
    "any": "ANY",
    "english": "EN",
    "french": "FR",
    "german": "DE",
    "dutch": "NL",
    "spanish": "ES",
    "bulgarian": "BG",
    "czech": "CS",
    "danish": "DA",
    "estonian": "ET",
    "greek": "EL",
    "irish": "GA",
    "croatian": "HR",
    "italian": "IT",
    "latvian": "LV",
    "lithuanian": "LT",
    "hungarian": "HU",
    "maltese": "MT",
    "polish": "PL",
    "portuguese": "PT",
    "romanian": "RO",
    "slovak": "SK",
    "slovenian": "SL",
    "finnish": "FI",
    "swedish": "SV",
    "icelandic": "IS",
    "norwegian": "NO",
    "turkish": "TR",
}


@dataclass
class SheetData:
    headers: list[str]
    values: list[object]

    def get(self, header: str, occurrence: int = 1) -> str:
        hits = [i for i, h in enumerate(self.headers) if (h or "").strip().lower() == header.strip().lower()]
        if len(hits) < occurrence:
            return ""
        value = self.values[hits[occurrence - 1]]
        if value is None:
            return ""
        return str(value).strip()


@dataclass
class SchemaConstraints:
    risk_classes: set[str]
    device_statuses: set[str]
    issuing_entities: set[str]
    certificate_types: set[str]
    basic_required_fields: set[str]
    udidi_required_fields: set[str]


class ExcelToEudamedXML:
    def __init__(self, workspace_root: Path):
        self.workspace_root = workspace_root

    def _resolve(self, path: str) -> Path:
        resolved = Path(path)
        return resolved if resolved.is_absolute() else (self.workspace_root / resolved)

    @staticmethod
    def _to_bool(value: str | object) -> str:
        raw = str(value or "").strip().lower()
        if raw in {"y", "yes", "true", "1"}:
            return "true"
        if raw in {"n", "no", "false", "0"}:
            return "false"
        return ""

    @staticmethod
    def _normalize_risk_class(value: str) -> str:
        raw = (value or "").strip().lower()
        return RISK_CLASS_MAP.get(raw, (value or "").strip().replace(" ", "_").upper())

    @staticmethod
    def _normalize_status(value: str) -> str:
        raw = (value or "").strip().lower()
        return STATUS_MAP.get(raw, (value or "").strip().replace(" ", "_").upper())

    @staticmethod
    def _normalize_language(value: str, default: str = "ANY") -> str:
        raw = (value or "").strip()
        if not raw:
            return default
        upper_raw = raw.upper()
        if len(upper_raw) in {2, 3}:
            return upper_raw
        return LANGUAGE_MAP.get(raw.lower(), default)

    @staticmethod
    def _normalize_certificate_type(value: str) -> str:
        raw = (value or "").strip()
        return raw.upper().replace("-", "_").replace(" ", "_")

    @staticmethod
    def _normalize_nb_actor_code(value: str) -> str:
        raw = (value or "").strip()
        digits = "".join(character for character in raw if character.isdigit())
        if len(digits) == 4:
            return digits
        return raw

    @staticmethod
    def _sheet_data(ws: openpyxl.worksheet.worksheet.Worksheet) -> SheetData:
        headers = [str(ws.cell(1, c).value).strip() if ws.cell(1, c).value is not None else "" for c in range(1, ws.max_column + 1)]
        values = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)] if ws.max_row >= 2 else []
        return SheetData(headers=headers, values=values)

    def _sheet_by_prefix(self, wb: openpyxl.Workbook, prefix: str) -> openpyxl.worksheet.worksheet.Worksheet:
        for name in wb.sheetnames:
            if name.lower().startswith(prefix.lower()):
                return wb[name]
        raise ValueError(f"Required sheet starting with '{prefix}' not found")

    def _optional_sheet_by_prefix(self, wb: openpyxl.Workbook, prefix: str) -> openpyxl.worksheet.worksheet.Worksheet | None:
        for name in wb.sheetnames:
            if name.lower().startswith(prefix.lower()):
                return wb[name]
        return None

    @staticmethod
    def _ns_tag(prefix: str, local_name: str) -> str:
        return f"{{{NS[prefix]}}}{local_name}"

    def _append(self, parent: etree._Element, prefix: str, local_name: str, text: str | None = None) -> etree._Element:
        element = etree.SubElement(parent, self._ns_tag(prefix, local_name))
        if text is not None:
            element.text = text
        return element

    def _append_bool(self, parent: etree._Element, prefix: str, local_name: str, raw_value: str, default: str = "false") -> etree._Element:
        value = self._to_bool(raw_value)
        if value == "":
            value = default
        return self._append(parent, prefix, local_name, value)

    def _append_required_text(self, parent: etree._Element, prefix: str, local_name: str, value: str, error_message: str) -> etree._Element:
        if not value:
            raise ValueError(error_message)
        return self._append(parent, prefix, local_name, value)

    def _append_text_or_comment(
        self,
        parent: etree._Element,
        prefix: str,
        local_name: str,
        value: str,
        comment_if_blank: str,
    ) -> etree._Element:
        text = (value or "").strip()
        if text:
            return self._append(parent, prefix, local_name, text)
        parent.append(etree.Comment(comment_if_blank))
        return self._append(parent, prefix, local_name, "")

    def _append_text_with_default_comment(
        self,
        parent: etree._Element,
        prefix: str,
        local_name: str,
        value: str,
        default_value: str,
        comment_if_defaulted: str,
    ) -> etree._Element:
        text = (value or "").strip()
        if text:
            return self._append(parent, prefix, local_name, text)
        parent.append(etree.Comment(comment_if_defaulted))
        return self._append(parent, prefix, local_name, default_value)

    @staticmethod
    def _extract_simple_type_enums(xsd_file: Path, type_name: str) -> set[str]:
        if not xsd_file.exists():
            return set()

        doc = etree.parse(str(xsd_file))
        xs_ns = {"xs": "http://www.w3.org/2001/XMLSchema"}
        values = doc.xpath(
            f"//xs:simpleType[@name='{type_name}']/xs:restriction/xs:enumeration/@value",
            namespaces=xs_ns,
        )
        return {str(value).strip() for value in values if str(value).strip()}

    @staticmethod
    def _type_local_name(type_name: str) -> str:
        if not type_name:
            return ""
        return type_name.split(":", 1)[-1]

    @staticmethod
    def _min_occurs(element: etree._Element) -> int:
        value = element.get("minOccurs")
        if value is None:
            return 1
        try:
            return int(value)
        except ValueError:
            return 1

    def _find_complex_type(self, docs: list[etree._ElementTree], type_name: str) -> etree._Element | None:
        local = self._type_local_name(type_name)
        xs_ns = {"xs": "http://www.w3.org/2001/XMLSchema"}
        for doc in docs:
            found = doc.xpath(f"//xs:complexType[@name='{local}']", namespaces=xs_ns)
            if found:
                return found[0]
        return None

    def _find_group(self, docs: list[etree._ElementTree], group_name: str) -> etree._Element | None:
        local = self._type_local_name(group_name)
        xs_ns = {"xs": "http://www.w3.org/2001/XMLSchema"}
        for doc in docs:
            found = doc.xpath(f"//xs:group[@name='{local}']", namespaces=xs_ns)
            if found:
                return found[0]
        return None

    def _collect_required_from_sequence(self, seq: etree._Element, docs: list[etree._ElementTree], out: set[str]) -> None:
        xs_ns = {"xs": "http://www.w3.org/2001/XMLSchema"}
        for child in seq:
            tag = etree.QName(child.tag).localname
            if tag == "element" and self._min_occurs(child) > 0:
                name = child.get("name")
                if name:
                    out.add(name)
            elif tag == "group" and self._min_occurs(child) > 0:
                ref = child.get("ref")
                if not ref:
                    continue
                group = self._find_group(docs, ref)
                if group is None:
                    continue
                group_seq = group.find("xs:sequence", namespaces=xs_ns)
                if group_seq is not None:
                    self._collect_required_from_sequence(group_seq, docs, out)

    def _collect_required_fields_for_type(self, docs: list[etree._ElementTree], type_name: str, visited: set[str] | None = None) -> set[str]:
        if visited is None:
            visited = set()
        local = self._type_local_name(type_name)
        if local in visited:
            return set()
        visited.add(local)

        required: set[str] = set()
        xs_ns = {"xs": "http://www.w3.org/2001/XMLSchema"}
        complex_type = self._find_complex_type(docs, local)
        if complex_type is None:
            return required

        extension = complex_type.find("xs:complexContent/xs:extension", namespaces=xs_ns)
        if extension is not None:
            base = extension.get("base")
            if base:
                required |= self._collect_required_fields_for_type(docs, base, visited)
            seq = extension.find("xs:sequence", namespaces=xs_ns)
            if seq is not None:
                self._collect_required_from_sequence(seq, docs, required)
            return required

        seq = complex_type.find("xs:sequence", namespaces=xs_ns)
        if seq is not None:
            self._collect_required_from_sequence(seq, docs, required)
        return required

    def _required_field_sets(self, xsd_root: Path) -> tuple[set[str], set[str]]:
        docs = [
            etree.parse(str(xsd_root / "data/Entity/Device/RegulationDevice/BasicUDIType.xsd")),
            etree.parse(str(xsd_root / "data/Entity/Device/RegulationDevice/UDIDIType.xsd")),
            etree.parse(str(xsd_root / "data/Entity/Device/CommonDeviceType.xsd")),
            etree.parse(str(xsd_root / "data/Entity/Entity.xsd")),
            etree.parse(str(xsd_root / "data/Entity/Links/LinkType.xsd")),
        ]

        basic_required = self._collect_required_fields_for_type(docs, "MDRBasicUDIType")
        udidi_required = self._collect_required_fields_for_type(docs, "MDRUDIDIDataType")

        # Expand required nested fields for complex types used in required elements.
        if "identifier" in basic_required:
            basic_required.add("identifier.DICode")
            basic_required.add("identifier.issuingEntityCode")
        if "identifier" in udidi_required:
            udidi_required.add("identifier.DICode")
            udidi_required.add("identifier.issuingEntityCode")
        if "basicUDIIdentifier" in udidi_required:
            udidi_required.add("basicUDIIdentifier.DICode")
            udidi_required.add("basicUDIIdentifier.issuingEntityCode")
        if "status" in udidi_required:
            udidi_required.add("status.code")

        return basic_required, udidi_required

    def _load_schema_constraints(self, xsd_file: Path) -> SchemaConstraints:
        xsd_root = xsd_file.parent.parent
        risk_file = xsd_root / "data/Entity/Common/RiskClassEnum.xsd"
        udidi_file = xsd_root / "data/Entity/Device/RegulationDevice/UDIDIType.xsd"
        link_file = xsd_root / "data/Entity/Links/LinkType.xsd"
        basic_required, udidi_required = self._required_field_sets(xsd_root)

        return SchemaConstraints(
            risk_classes=self._extract_simple_type_enums(risk_file, "RiskClassEnum"),
            device_statuses=self._extract_simple_type_enums(udidi_file, "DeviceStatusEnum"),
            issuing_entities=self._extract_simple_type_enums(udidi_file, "IssuingEntityTypeEnum"),
            certificate_types=self._extract_simple_type_enums(link_file, "GenericCertificateTypeEnum"),
            basic_required_fields=basic_required,
            udidi_required_fields=udidi_required,
        )

    def _build_source_values(self, basic: SheetData, ident: SheetData, chars: SheetData, device: SheetData) -> dict[str, str]:
        status = self._normalize_status(ident.get("UDI-DI status ( On the EU market / No longer placed on the EU market / Not intended for the EU market )"))
        return {
            "riskClass": self._normalize_risk_class(basic.get("Risk class*")),
            "model": basic.get("Device Model"),
            "modelName": basic.get("Device Name*"),
            "identifier.DICode": basic.get("Basic UDI-DI code") or basic.get("Basic UDI code") or ident.get("UDI-DI code"),
            "identifier.issuingEntityCode": basic.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)") or ident.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)"),
            "animalTissuesCells": self._to_bool(device.get("Tissues and cells Presence of animal tissues or cells, or their derivatives:")) or "false",
            "humanTissuesCells": self._to_bool(device.get("Tissues and cells Presence of human tissues or cells, or their derivatives:")) or "false",
            "MFActorCode": basic.get("Manufacturer Actor Code (SRN)") or basic.get("MFActorCode"),
            "humanProductCheck": self._to_bool(device.get("Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma(Y/N)")) or "false",
            "medicinalProductCheck": self._to_bool(device.get("Information on substances Presence of a substance which, if used separately, may be considered to be a medicinal product (Y/N)")) or "false",
            "type": "DEVICE",
            "active": self._to_bool(basic.get("Active device*")) or "false",
            "administeringMedicine": self._to_bool(basic.get("Device intended to administer and/or remove medicinal product*")) or "false",
            "implantable": self._to_bool(basic.get("Implantable*")) or "false",
            "measuringFunction": self._to_bool(basic.get("Measuring function*")) or "false",
            "reusable": self._to_bool(basic.get("Reusable surgical instrument*")) or "false",
            "status.code": status,
            "basicUDIIdentifier.DICode": basic.get("Basic UDI-DI code") or basic.get("Basic UDI code") or ident.get("UDI-DI code"),
            "basicUDIIdentifier.issuingEntityCode": basic.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)") or ident.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)"),
            "MDNCodes": ident.get("Enter a nomenclature code (EMDN code)"),
            "referenceNumber": ident.get("Reference/Catalogue number") or ident.get("UDI-DI code"),
            "sterile": self._to_bool(chars.get("Device labelled as sterile")) or "false",
            "sterilization": self._to_bool(chars.get("Need for sterilisation before use")) or "false",
            "numberOfReuses": (
                "0"
                if self._to_bool(chars.get("Labelled as single use")) == "true"
                else (str(chars.get("Maximum number of reuses")) if str(chars.get("Maximum number of reuses") or "").isdigit() else "1")
            ),
            "latex": self._to_bool(chars.get("Containing latex")) or "false",
            "reprocessed": self._to_bool(device.get("Reprocessed single use device")) or "false",
            # New conditional business rule fields
            "isSystemSPP": self._to_bool(basic.get("Is it a System or Procedure pack which is a Device in itself (Y/N)")) or "false",
            "isKit": self._to_bool(basic.get("Is it a Kit (Y/N)")) or "false",
            "specialDeviceType": basic.get("Special Device Type (if applicable)") or "",
            "iibImplantableExceptions": self._to_bool(basic.get("IIb implantable exceptions: Is device a suture/staple/dental/screw/etc (Y/N)")) or "false",
            "iibImplantableDeviceType": basic.get("IIb implantable exceptions: Specify device type (suture, staple, dental filling, screw, etc.)") or "",
            "memberStatePlacement": ident.get("Member State where device is placed on market (e.g., IT, DE, FR)") or "",
            "memberStatesMadeAvailable": ident.get("Member States where device is made available (comma-separated, e.g., IT,DE,FR)") or "",
            "additionalDescription": ident.get("Additional product description") or "",
        }

    def _validate_input_data(self, basic: SheetData, ident: SheetData, chars: SheetData, device: SheetData, cert: SheetData, constraints: SchemaConstraints) -> None:
        errors: list[str] = []
        source_values = self._build_source_values(basic, ident, chars, device)

        risk_raw = basic.get("Risk class*")
        risk_class = self._normalize_risk_class(risk_raw)
        if not risk_class:
            errors.append("'Risk class*' is required.")
        elif constraints.risk_classes and risk_class not in constraints.risk_classes:
            allowed = ", ".join(sorted(constraints.risk_classes))
            errors.append(f"Risk class '{risk_raw}' normalized to '{risk_class}' is not valid per XSD. Allowed: {allowed}")

        if not (basic.get("Device Model") or basic.get("Device Name*")):
            errors.append("Either 'Device Model' or 'Device Name*' must be provided.")

        issuing_entity = ident.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)")
        if not issuing_entity:
            errors.append("'UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)' is required.")
        elif constraints.issuing_entities and issuing_entity.strip().upper() not in constraints.issuing_entities:
            allowed = ", ".join(sorted(constraints.issuing_entities))
            errors.append(f"Issuing Entity '{issuing_entity}' is not valid per XSD. Allowed: {allowed}")

        if not ident.get("UDI-DI code"):
            errors.append("'UDI-DI code' is required.")

        status_raw = ident.get("UDI-DI status ( On the EU market / No longer placed on the EU market / Not intended for the EU market )")
        status = self._normalize_status(status_raw)
        if not status:
            errors.append("'UDI-DI status (...)' is required.")
        elif constraints.device_statuses and status not in constraints.device_statuses:
            allowed = ", ".join(sorted(constraints.device_statuses))
            errors.append(f"UDI-DI status '{status_raw}' normalized to '{status}' is not valid per XSD. Allowed: {allowed}")

        cert_applicable = self._to_bool(cert.get("Certificate applicable"))
        if risk_class == "CLASS_III" and cert_applicable != "true":
            errors.append("Certificate information is required for Class III devices. Set 'Certificate applicable' to Yes.")

        if cert_applicable == "true":
            nb_actor_code = self._normalize_nb_actor_code(cert.get("Notified Body Actor Code (NBActorCode)"))
            if not nb_actor_code:
                errors.append("'Notified Body Actor Code (NBActorCode)' is required when certificate is applicable.")
            elif not (len(nb_actor_code) == 4 and nb_actor_code.isdigit()):
                errors.append("NBActorCode must be exactly 4 digits as per XSD pattern [0-9]{4}.")

            cert_type = self._normalize_certificate_type(cert.get("Certificate type"))
            if not cert_type:
                errors.append("'Certificate type' is required when certificate is applicable.")
            elif constraints.certificate_types and cert_type not in constraints.certificate_types:
                allowed = ", ".join(sorted(constraints.certificate_types))
                errors.append(f"Certificate type '{cert_type}' is not valid per XSD. Allowed: {allowed}")

        # Dynamic required-field checks driven by XSD minOccurs.
        basic_required = set(constraints.basic_required_fields)
        udidi_required = set(constraints.udidi_required_fields)

        # Container complex elements are validated through their required child fields.
        basic_required.discard("identifier")
        udidi_required.discard("identifier")
        udidi_required.discard("basicUDIIdentifier")
        udidi_required.discard("status")

        # These are intentionally allowed to be blank so output can contain user prompts.
        basic_required.discard("MFActorCode")
        udidi_required.discard("MDNCodes")
        udidi_required.discard("referenceNumber")

        # Choice in BasicUDIType is model OR modelName; treat as one logical requirement.
        if "model" in basic_required or "modelName" in basic_required:
            if not (source_values.get("model") or source_values.get("modelName")):
                errors.append("XSD requires one of 'model' or 'modelName'; provide Device Model or Device Name*.")
            basic_required.discard("model")
            basic_required.discard("modelName")

        for field in sorted(basic_required):
            value = source_values.get(field, "")
            if value in {"", None}:
                errors.append(f"Missing required field for MDRBasicUDIType (from XSD): {field}")

        for field in sorted(udidi_required):
            value = source_values.get(field, "")
            if value in {"", None}:
                errors.append(f"Missing required field for MDRUDIDIDataType (from XSD): {field}")

        # ===== Additional Conditional Business Rules (BR-UDID) =====
        # BR-UDID-676: Class I → Implantable forced to False
        if risk_class == "CLASS_I":
            implantable = self._to_bool(basic.get("Implantable*"))
            if implantable == "true":
                errors.append("BR-UDID-676: Class I devices cannot be implantable. Auto-correcting to No.")

        # BR-UDID-677: Implantable → Reusable surgical instrument forced to False
        implantable = self._to_bool(basic.get("Implantable*"))
        if implantable == "true":
            reusable = self._to_bool(basic.get("Reusable surgical instrument*"))
            if reusable == "true":
                errors.append("BR-UDID-677: Implantable devices cannot be reusable surgical instruments. Auto-correcting to No.")

        # BR-UDID-024: Max reuses only if NOT single-use
        single_use = self._to_bool(chars.get("Labelled as single use"))
        if single_use == "true":
            max_reuses_raw = chars.get("Maximum number of reuses")
            if max_reuses_raw and str(max_reuses_raw).strip() and int(str(max_reuses_raw).strip() or "0") > 0:
                errors.append("BR-UDID-024: Maximum number of reuses cannot be provided for single-use devices.")

        # BR-UDID-043: Member States required for Class IIa/IIb/III when on market
        if risk_class in {"CLASS_IIA", "CLASS_IIB", "CLASS_III"} and status == "ON_THE_MARKET":
            member_state_placement = ident.get("Member State where device is placed on market (e.g., IT, DE, FR)")
            member_states_available = ident.get("Member States where device is made available (comma-separated, e.g., IT,DE,FR)")
            if not member_state_placement:
                errors.append("BR-UDID-043 / BR-UDID-045: Member State for placement on market is required for Class IIa/IIb/III devices on market.")
            if not member_states_available:
                errors.append("BR-UDID-043 / BR-UDID-673 / BR-UDID-674: Member States where device is made available are required for Class IIa/IIb/III devices on market.")

        # BR-UDID-113: Directive Certificates (conditional by risk class and legislation)
        # For now: Class III requires certificate; note on optional cases
        # In legacy MDD Class I (non-measuring) and IVDD General, certificates are optional
        if risk_class == "CLASS_III" and cert_applicable != "true":
            errors.append("BR-UDID-113: Class III devices require certificate information. Set 'Certificate applicable' to Yes.")

        # BR-UDID-705: Special Device Type cannot be set with Kit or System/SPP
        special_type = basic.get("Special Device Type (if applicable)")
        is_system_spp = self._to_bool(basic.get("Is it a System or Procedure pack which is a Device in itself (Y/N)"))
        is_kit = self._to_bool(basic.get("Is it a Kit (Y/N)"))
        if special_type and special_type.strip():
            if is_system_spp == "true" or is_kit == "true":
                errors.append("BR-UDID-705: Special Device Type cannot be set if the device is marked as Kit or System/Procedure pack.")

        # BR-UDID-131: Additional description mandatory for System/SPP/Kit
        additional_description = ident.get("Additional product description")
        if is_system_spp == "true" or is_kit == "true":
            if not additional_description or not additional_description.strip():
                errors.append("BR-UDID-131: Additional product description is mandatory for System, Procedure pack, or Kit devices.")

        # BR-UDID-635: Class IIb implantable → device type specification required
        if risk_class == "CLASS_IIB" and implantable == "true":
            device_type_checkbox = self._to_bool(basic.get("IIb implantable exceptions: Is device a suture/staple/dental/screw/etc (Y/N)"))
            device_type_spec = basic.get("IIb implantable exceptions: Specify device type (suture, staple, dental filling, screw, etc.)")
            if device_type_checkbox == "true":
                if not device_type_spec or not device_type_spec.strip():
                    errors.append("BR-UDID-635: For Class IIb implantable devices, the device type must be specified (suture, staple, dental filling, dental brace, tooth crown, screw, wedge, plate, wire, pin, clip, or connector).")
                # Validate it's one of the allowed types
                allowed_types = {"suture", "staple", "dental filling", "dental brace", "tooth crown", "screw", "wedge", "plate", "wire", "pin", "clip", "connector"}
                device_type_normalized = device_type_spec.strip().lower()
                if device_type_normalized not in allowed_types:
                    errors.append(f"BR-UDID-635: Device type '{device_type_spec}' is not in the allowed list. Must be one of: {', '.join(sorted(allowed_types))}.")

        if errors:
            raise ValueError("Input validation failed:\n- " + "\n- ".join(errors))

    def _device_node(self, root: etree._Element) -> etree._Element:
        node = root.find(f".//{{{NS['device']}}}Device")
        if node is None:
            raise ValueError("Device payload node was not created.")
        return node

    def _build_message_root(
        self,
        recipient_node_actor_code: str,
        sender_node_actor_code: str,
        conversation_id: str,
        correlation_id: str,
        message_id: str,
    ) -> etree._Element:
        # Caller must provide envelope IDs and node actor codes explicitly.
        root = etree.Element(self._ns_tag("message", "PullResponse"), nsmap=XML_NS_MAP)
        root.set("version", "3.0.28")

        self._append_text_with_default_comment(
            root,
            "message",
            "conversationID",
            conversation_id,
            DEFAULT_CONVERSATION_ID,
            "Provide your own conversationID (default value applied)",
        )
        self._append_text_with_default_comment(
            root,
            "message",
            "correlationID",
            correlation_id,
            DEFAULT_CORRELATION_ID,
            "Provide your own correlationID (default value applied)",
        )
        self._append(root, "message", "creationDateTime", datetime.now(timezone.utc).isoformat(timespec="milliseconds"))
        self._append_text_with_default_comment(
            root,
            "message",
            "messageID",
            message_id,
            DEFAULT_MESSAGE_ID,
            "Provide your own messageID (default value applied)",
        )

        recipient = self._append(root, "message", "recipient")
        recipient_node = self._append(recipient, "message", "node")
        self._append_text_with_default_comment(
            recipient_node,
            "service",
            "nodeActorCode",
            recipient_node_actor_code,
            DEFAULT_RECIPIENT_NODE_ACTOR_CODE,
            "Provide your own recipient nodeActorCode (default value applied)",
        )
        recipient_service = self._append(recipient, "message", "service")
        self._append(recipient_service, "service", "serviceID", "DEVICE")
        self._append(recipient_service, "service", "serviceOperation", "GET")

        payload = self._append(root, "message", "payload")
        device = etree.SubElement(payload, self._ns_tag("device", "Device"))
        device.set(self._ns_tag("xsi", "type"), "device:MDRDeviceType")

        sender = self._append(root, "message", "sender")
        sender_node = self._append(sender, "message", "node")
        self._append_text_with_default_comment(
            sender_node,
            "service",
            "nodeActorCode",
            sender_node_actor_code,
            DEFAULT_SENDER_NODE_ACTOR_CODE,
            "Provide your own sender nodeActorCode (default value applied)",
        )
        sender_service = self._append(sender, "message", "service")
        self._append(sender_service, "service", "serviceID", "DEVICE")
        self._append(sender_service, "service", "serviceOperation", "GET")

        self._append(root, "message", "numberOfPages", "1")
        self._append(root, "message", "pageNumber", "0")
        self._append(root, "message", "pageSize", "20")
        self._append(root, "message", "responseCode", "SUCCESS")
        return root

    def _basic_udi_model_choice(self, parent: etree._Element, basic: SheetData) -> None:
        model_applicable = self._to_bool(basic.get("Device model applicable"))
        model_value = basic.get("Device Model")
        device_name = basic.get("Device Name*")

        if model_value and device_name:
            model_name = self._append(parent, "budi", "modelName")
            self._append(model_name, "commondevice", "model", model_value)
            self._append(model_name, "commondevice", "name", device_name)
            return

        if model_value and model_applicable != "false":
            self._append(parent, "budi", "model", model_value)
            return

        if device_name:
            model_name = self._append(parent, "budi", "modelName")
            self._append(model_name, "commondevice", "name", device_name)
            return

        if model_value:
            self._append(parent, "budi", "model", model_value)
            return

        raise ValueError("Either 'Device Model' or 'Device Name*' must be provided.")

    def _append_basic_udi_identifier(self, parent: etree._Element, basic: SheetData, ident: SheetData) -> None:
        basic_identifier_code = basic.get("Basic UDI-DI code") or basic.get("Basic UDI code") or ident.get("UDI-DI code")
        issuing_entity = basic.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)") or ident.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)")

        identifier = self._append(parent, "budi", "identifier")
        self._append_required_text(identifier, "commondevice", "DICode", basic_identifier_code, "Basic UDI-DI code is required.")
        self._append_required_text(identifier, "commondevice", "issuingEntityCode", issuing_entity, "Issuing entity is required for the Basic UDI-DI.")

    def _append_certificate_links(self, parent: etree._Element, ident: SheetData, cert: SheetData, risk_class: str) -> None:
        cert_applicable = self._to_bool(cert.get("Certificate applicable"))
        if risk_class == "CLASS_III" and cert_applicable != "true":
            raise ValueError("Certificate information is required for Class III devices. Set 'Certificate applicable' to Yes.")

        if cert_applicable != "true":
            return

        nb_actor_code = cert.get("Notified Body Actor Code (NBActorCode)")
        certificate_type = cert.get("Certificate type")
        certificate_number = cert.get("Certificate number") or ident.get("UDI-DI code")

        if not nb_actor_code:
            raise ValueError("'Notified Body Actor Code (NBActorCode)' is required when certificate is applicable.")
        if not certificate_type:
            raise ValueError("'Certificate type' is required when certificate is applicable.")

        links_parent = self._append(parent, "budi", "deviceCertificateLinks")
        link = self._append(links_parent, "lnks", "deviceCertificateLink")
        self._append(link, "lnks", "certificateNumber", certificate_number)
        expiry_date = cert.get("Certificate expiry date (YYYY-MM-DD)")
        if expiry_date:
            self._append(link, "lnks", "expiryDate", expiry_date)

        self._append(link, "lnks", "NBActorCode", self._normalize_nb_actor_code(nb_actor_code))
        revision_number = cert.get("Certificate revision number")
        if revision_number:
            self._append(link, "lnks", "certificateRevisionNumber", revision_number)

        self._append(link, "lnks", "certificateType", self._normalize_certificate_type(certificate_type))

    def _build_basic_udi(self, root: etree._Element, basic: SheetData, device: SheetData, ident: SheetData, cert: SheetData) -> etree._Element:
        device_node = self._device_node(root)
        basic_udi = self._append(device_node, "device", "MDRBasicUDI")

        self._append(basic_udi, "e", "state", "REGISTERED")
        self._append(basic_udi, "e", "version", "1")
        self._append(basic_udi, "e", "versionDate", datetime.now(timezone.utc).isoformat(timespec="milliseconds"))

        risk_class = self._normalize_risk_class(basic.get("Risk class*"))
        if not risk_class:
            raise ValueError("'Risk class*' is required.")
        self._append(basic_udi, "budi", "riskClass", risk_class)

        self._basic_udi_model_choice(basic_udi, basic)
        self._append_basic_udi_identifier(basic_udi, basic, ident)

        human_tissues = device.get("Tissues and cells Presence of human tissues or cells, or their derivatives:")
        animal_tissues = device.get("Tissues and cells Presence of animal tissues or cells, or their derivatives:")
        medicinal_product = device.get("Information on substances Presence of a substance which, if used separately, may be considered to be a medicinal product (Y/N)")
        human_blood_product = device.get("Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma(Y/N)")

        self._append_bool(basic_udi, "budi", "animalTissuesCells", animal_tissues, default="false")
        self._append_bool(basic_udi, "budi", "humanTissuesCells", human_tissues, default="false")

        manufacturer_srn = basic.get("Manufacturer Actor Code (SRN)") or basic.get("MFActorCode")
        self._append_text_with_default_comment(
            basic_udi,
            "budi",
            "MFActorCode",
            manufacturer_srn,
            DEFAULT_MF_ACTOR_CODE,
            "Provide your own MFActorCode (Manufacturer Actor Code / SRN) (default value applied)",
        )

        self._append_certificate_links(basic_udi, ident, cert, risk_class)

        self._append_bool(basic_udi, "budi", "humanProductCheck", human_blood_product, default="false")
        self._append_bool(basic_udi, "budi", "IIb_implantable_exceptions", basic.get("IIb implantable exceptions"), default="false")
        self._append_bool(basic_udi, "budi", "medicinalProductCheck", medicinal_product, default="false")

        self._append(basic_udi, "budi", "type", "DEVICE")

        self._append_bool(basic_udi, "commondevice", "active", basic.get("Active device*"), default="false")
        self._append_bool(basic_udi, "commondevice", "administeringMedicine", basic.get("Device intended to administer and/or remove medicinal product*"), default="false")
        self._append_bool(basic_udi, "commondevice", "implantable", basic.get("Implantable*"), default="false")
        self._append_bool(basic_udi, "commondevice", "measuringFunction", basic.get("Measuring function*"), default="false")
        self._append_bool(basic_udi, "commondevice", "reusable", basic.get("Reusable surgical instrument*"), default="false")

        return basic_udi

    def _build_udidi_data(self, root: etree._Element, ident: SheetData, chars: SheetData, device: SheetData, basic: SheetData) -> etree._Element:
        device_node = self._device_node(root)
        udidi_data = self._append(device_node, "device", "MDRUDIDIData")

        self._append(udidi_data, "e", "state", "REGISTERED")
        self._append(udidi_data, "e", "version", "1")
        self._append(udidi_data, "e", "versionDate", datetime.now(timezone.utc).isoformat(timespec="milliseconds"))

        identifier = self._append(udidi_data, "udidi", "identifier")
        self._append_required_text(identifier, "commondevice", "DICode", ident.get("UDI-DI code"), "UDI-DI code is required.")
        self._append_required_text(identifier, "commondevice", "issuingEntityCode", ident.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)"), "UDI-DI issuing entity is required.")

        status = self._append(udidi_data, "udidi", "status")
        self._append_text_with_default_comment(
            status,
            "commondevice",
            "code",
            self._normalize_status(ident.get("UDI-DI status ( On the EU market / No longer placed on the EU market / Not intended for the EU market )")),
            DEFAULT_UDI_STATUS_CODE,
            "Provide your own UDI-DI status code (default value applied)",
        )

        basic_identifier = self._append(udidi_data, "udidi", "basicUDIIdentifier")
        basic_code = basic.get("Basic UDI-DI code") or basic.get("Basic UDI code") or ident.get("UDI-DI code")
        basic_issuer = basic.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)") or ident.get("UDI-DI identification Issuing Entity (GS1/HIBCC/ICCBBA/IFA)")
        self._append_required_text(basic_identifier, "commondevice", "DICode", basic_code, "Basic UDI-DI code is required.")
        self._append_required_text(basic_identifier, "commondevice", "issuingEntityCode", basic_issuer, "Basic UDI-DI issuing entity is required.")

        self._append_text_with_default_comment(
            udidi_data,
            "udidi",
            "MDNCodes",
            ident.get("Enter a nomenclature code (EMDN code)"),
            DEFAULT_MDN_CODE,
            "Provide your own EMDN code for MDNCodes (default value applied)",
        )

        production_tokens: list[str] = []
        if self._to_bool(ident.get("Type of UDI-PI Lot or Batch number")) == "true":
            production_tokens.append("BATCH_NUMBER")
        if self._to_bool(ident.get("Type of UDI-PI Serial number")) == "true":
            production_tokens.append("SERIALISATION_NUMBER")
        if self._to_bool(ident.get("Type of UDI-PI Manufacturing date")) == "true":
            production_tokens.append("MANUFACTURING_DATE")
        if self._to_bool(ident.get("Type of UDI-PI Expiration date")) == "true":
            production_tokens.append("EXPIRATION_DATE")
        if production_tokens:
            self._append(udidi_data, "udidi", "productionIdentifier", " ".join(production_tokens))

        self._append_text_with_default_comment(
            udidi_data,
            "udidi",
            "referenceNumber",
            ident.get("Reference/Catalogue number") or ident.get("UDI-DI code"),
            DEFAULT_REFERENCE_NUMBER,
            "Provide your own reference/catalogue number (default value applied)",
        )

        secondary_applicable = self._to_bool(ident.get("UDI-DI from another entity (secondary) applicable"))
        sec_issuer = ident.get("Secondary Issuing Entity (GS1/HIBCC/ICCBBA/IFA)")
        sec_code = ident.get("Secondary UDI-DI code")
        if secondary_applicable == "true" and sec_issuer and sec_code:
            secondary = self._append(udidi_data, "udidi", "secondaryIdentifier")
            self._append_required_text(secondary, "commondevice", "DICode", sec_code, "Secondary UDI-DI code is required.")
            self._append_required_text(secondary, "commondevice", "issuingEntityCode", sec_issuer, "Secondary UDI-DI issuer is required.")

        self._append_bool(udidi_data, "udidi", "sterile", chars.get("Device labelled as sterile"), default="false")
        self._append_bool(udidi_data, "udidi", "sterilization", chars.get("Need for sterilisation before use"), default="false")

        trade_name = ident.get("Trade name")
        if trade_name:
            trade_names = self._append(udidi_data, "udidi", "tradeNames")
            trade_name_entry = self._append(trade_names, "lngs", "name")
            self._append(trade_name_entry, "lngs", "language", self._normalize_language(ident.get("Select the language"), default="ANY"))
            self._append(trade_name_entry, "lngs", "textValue", trade_name)

        additional_description = ident.get("Additional product description")
        if additional_description:
            additional = self._append(udidi_data, "udidi", "additionalDescription")
            additional_entry = self._append(additional, "lngs", "name")
            self._append(additional_entry, "lngs", "language", self._normalize_language(ident.get("Select the language", occurrence=2), default="ANY"))
            self._append(additional_entry, "lngs", "textValue", additional_description)

        website = ident.get("URL for additional information (as electronic instructions for use)")
        if website:
            self._append(udidi_data, "udidi", "website", website)

        single_use = self._to_bool(chars.get("Labelled as single use"))
        max_reuses = chars.get("Maximum number of reuses")
        if single_use == "true":
            self._append(udidi_data, "udidi", "numberOfReuses", "0")
        elif max_reuses and str(max_reuses).isdigit():
            self._append(udidi_data, "udidi", "numberOfReuses", str(max_reuses))
        else:
            self._append(udidi_data, "udidi", "numberOfReuses", "1")

        self._append(udidi_data, "udidi", "baseQuantity", ident.get("Quantity of device") or "1")
        self._append_bool(udidi_data, "udidi", "latex", chars.get("Containing latex"), default="false")
        self._append_bool(udidi_data, "udidi", "reprocessed", device.get("Reprocessed single use device"), default="false")

        if self._to_bool(chars.get("Clinical size application status")) == "true":
            pass

        return udidi_data

    def build_xml_from_excel(
        self,
        excel_file: Path,
        xsd_file: Path,
        recipient_node_actor_code: str,
        sender_node_actor_code: str,
        conversation_id: str,
        correlation_id: str,
        message_id: str,
    ) -> etree._ElementTree:
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        basic_ws = self._sheet_by_prefix(wb, "Basic UDI-DI information")
        cert_ws = self._optional_sheet_by_prefix(wb, "Certificate information")
        ident_ws = self._sheet_by_prefix(wb, "UDI-DI identification")
        chars_ws = self._sheet_by_prefix(wb, "UDI-DI characteristics")
        device_ws = self._sheet_by_prefix(wb, "Device information")

        basic = self._sheet_data(basic_ws)
        cert = self._sheet_data(cert_ws) if cert_ws is not None else SheetData(headers=[], values=[])
        ident = self._sheet_data(ident_ws)
        chars = self._sheet_data(chars_ws)
        device = self._sheet_data(device_ws)

        constraints = self._load_schema_constraints(xsd_file)
        self._validate_input_data(basic, ident, chars, device, cert, constraints)

        recipient_code = (recipient_node_actor_code or "").strip()
        sender_code = (sender_node_actor_code or "").strip()

        root = self._build_message_root(
            recipient_node_actor_code=recipient_code,
            sender_node_actor_code=sender_code,
            conversation_id=conversation_id,
            correlation_id=correlation_id,
            message_id=message_id,
        )
        self._build_basic_udi(root, basic, device, ident, cert)
        self._build_udidi_data(root, ident, chars, device, basic)
        return etree.ElementTree(root)

    def _validate_xsd(self, xml_file: Path, xsd_file: Path) -> None:
        if not xsd_file.exists():
            raise FileNotFoundError(f"XSD not found: {xsd_file}")

        original_cwd = os.getcwd()
        xsd_root = xsd_file.parent.parent.parent
        try:
            os.chdir(xsd_root)
            xsd_doc = etree.parse(str(xsd_file.relative_to(xsd_root)))
            schema = etree.XMLSchema(xsd_doc)
            xml_doc = etree.parse(str(xml_file))
            if not schema.validate(xml_doc):
                issues = [f"Line {error.line}, Col {error.column}: {error.message}" for error in schema.error_log]
                raise ValueError("XSD validation failed:\n" + "\n".join(issues))
            print("XSD validation: PASS")
        finally:
            os.chdir(original_cwd)

    def generate(
        self,
        excel_path: str,
        output_xml_path: str,
        xsd_path: str,
        recipient_node_actor_code: str,
        sender_node_actor_code: str,
        conversation_id: str,
        correlation_id: str,
        message_id: str,
    ) -> Path:
        excel_file = self._resolve(excel_path)
        output_file = self._resolve(output_xml_path)
        xsd_file = self._resolve(xsd_path)

        if not excel_file.exists():
            raise FileNotFoundError(f"Excel not found: {excel_file}")

        tree = self.build_xml_from_excel(
            excel_file,
            xsd_file,
            recipient_node_actor_code=recipient_node_actor_code,
            sender_node_actor_code=sender_node_actor_code,
            conversation_id=conversation_id,
            correlation_id=correlation_id,
            message_id=message_id,
        )
        output_file.parent.mkdir(parents=True, exist_ok=True)
        tree.write(str(output_file), pretty_print=True, xml_declaration=True, encoding="utf-8")
        self._validate_xsd(output_file, xsd_file)
        return output_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate EUDAMED XML from Excel and validate against XSD.")
    parser.add_argument("--excel", required=True, help="Input Excel path")
    parser.add_argument("--output", required=True, help="Output XML path")
    parser.add_argument(
        "--xsd",
        default="XSD_schemas/service/Message.xsd",
        help="Root XSD for validation (default: XSD_schemas/service/Message.xsd)",
    )
    parser.add_argument(
        "--recipient-node-actor-code",
        default="",
        help="message.recipient.node.nodeActorCode. Leave empty to emit blank value with XML comment.",
    )
    parser.add_argument(
        "--sender-node-actor-code",
        default="",
        help="message.sender.node.nodeActorCode. Leave empty to emit blank value with XML comment.",
    )
    parser.add_argument(
        "--conversation-id",
        default="",
        help="message.conversationID. Leave empty to emit blank value with XML comment.",
    )
    parser.add_argument(
        "--correlation-id",
        default="",
        help="message.correlationID. Leave empty to emit blank value with XML comment.",
    )
    parser.add_argument(
        "--message-id",
        default="",
        help="message.messageID. Leave empty to emit blank value with XML comment.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    workspace_root = Path(__file__).resolve().parent
    generator = ExcelToEudamedXML(workspace_root)
    output = generator.generate(
        excel_path=args.excel,
        output_xml_path=args.output,
        xsd_path=args.xsd,
        recipient_node_actor_code=args.recipient_node_actor_code,
        sender_node_actor_code=args.sender_node_actor_code,
        conversation_id=args.conversation_id,
        correlation_id=args.correlation_id,
        message_id=args.message_id,
    )
    print(f"Generated XML: {output}")


if __name__ == "__main__":
    main()
