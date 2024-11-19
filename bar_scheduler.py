import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, PatternFill, Font, Side
from datetime import datetime, timedelta
import random
from openpyxl.utils import get_column_letter


class BarScheduler:
    def __init__(self):
        self.YEAR = 2024
        self.MONTH = 11
        self.MONTH_NAMES = {
            1: "January",
            2: "February",
            3: "March",
            4: "April",
            5: "May",
            6: "June",
            7: "July",
            8: "August",
            9: "September",
            10: "October",
            11: "November",
            12: "December",
        }
        self.MONTH_NAME = self.MONTH_NAMES[self.MONTH]
        self.USERPATH = "/Users/martin/Desktop/"
        self.FILEPATH = self.USERPATH + f"{self.MONTH_NAME} (Svar) - Skjemasvar 1.csv"
        self.shifts = {
            "opening": {"color": "FFB4C6", "time": "12:30-17:00"},
            "middle": {"color": "B4D7FF", "time": "16:50-20:30"},
            "closing": {"color": "C6FFB4", "time": "20:20-00:30"},
        }
        self.weekend_color = "808080"
        self.no_reply_color = "404040"
        self.manual_review = []
        self.unmatched_availability = {}
        self.no_reply_members = set()

    def is_weekend(self, date_str):
        try:
            day = int(date_str.split(".")[0])
            return datetime(self.YEAR, self.MONTH, day).weekday() >= 5
        except ValueError:
            return True

    def is_monday(self, date_str):
        try:
            day = int(date_str.split(".")[0])
            return datetime(self.YEAR, self.MONTH, day).weekday() == 0
        except ValueError:
            return False

    def get_available_shifts(self, date):
        if self.is_monday(date):
            return ["opening", "middle"]
        return ["opening", "middle", "closing"]

    def get_next_weekend_dates(self, current_date, next_date):
        try:
            current_day = int(current_date.split(".")[0])
            next_day = int(next_date.split(".")[0])
            weekend_dates = []

            current = datetime(self.YEAR, self.MONTH, current_day)
            while current.weekday() < 5 and current.day < next_day:
                current = current + timedelta(days=1)

            if current.day < next_day:
                weekend_dates.append("WEEKEND")

            return weekend_dates
        except ValueError:
            return []

    def get_staff_requirement(self, date_str, shift_type):
        try:
            day = int(date_str.split(".")[0])
            weekday = datetime(self.YEAR, self.MONTH, day).weekday()

            requirements = {
                0: {"opening": 2, "middle": 2, "closing": 0},  # Monday
                1: {"opening": 2, "middle": 3, "closing": 3},  # Tuesday
                2: {"opening": 2, "middle": 3, "closing": 3},  # Wednesday
                3: {"opening": 2, "middle": 2, "closing": 2},  # Thursday
                4: {"opening": 2, "middle": 3, "closing": 3},  # Friday
            }

            return requirements.get(weekday, {}).get(shift_type, 2)
        except ValueError:
            return 2

    def find_member_match(self, input_name, member_list):
        def normalize_name(name):
            return "".join(c.lower() for c in name if c.isalnum())

        if input_name.lower() in (m.lower() for m in member_list):
            return next(m for m in member_list if m.lower() == input_name.lower())

        matches = []
        input_normalized = normalize_name(input_name)
        input_parts = set(normalize_name(p) for p in input_name.split())

        for member in member_list:
            member_normalized = normalize_name(member)
            member_parts = set(normalize_name(p) for p in member.split())

            score = 0
            if input_normalized == member_normalized:
                score = 1.0
            elif input_parts and member_parts:
                # Strong first name match
                if normalize_name(input_name.split()[0]) == normalize_name(
                    member.split()[0]
                ):
                    score = 0.95
                else:
                    # Part matching
                    common_parts = input_parts.intersection(member_parts)
                    score = (
                        len(common_parts) / len(input_parts.union(member_parts))
                        if common_parts
                        else 0
                    )

            if score > 0:
                matches.append((member, score))

        if matches:
            best_match = max(matches, key=lambda x: x[1])
            if best_match[1] >= 0.8:
                return best_match[0]
            elif best_match[1] >= 0.5:
                self.manual_review.append(
                    {
                        "input_name": input_name,
                        "possible_match": best_match[0],
                        "confidence": f"{best_match[1]:.2f}",
                    }
                )
                return best_match[0]

        self.manual_review.append(
            {
                "input_name": input_name,
                "possible_match": "No match found",
                "confidence": "0.00",
            }
        )
        return input_name

    def parse_shifts(self, cell_value, date_str):
        if (
            not isinstance(cell_value, str)
            or "Kan ikke jobbe denne dagen" in cell_value
        ):
            return []

        shifts = []
        shift_times = {
            "opening": "12:30-17:00",
            "middle": "16:50-20:30",
            "closing": "20:20-00:30",
        }

        for shift_type, time in shift_times.items():
            if time in cell_value:
                if not (self.is_monday(date_str) and shift_type == "closing"):
                    shifts.append(shift_type)
        return shifts

    def check_consecutive_days(self, schedule, staff_name, current_date, dates):
        current_idx = dates.index(current_date)
        if current_idx > 0:
            prev_date = dates[current_idx - 1]
            if not self.is_weekend(prev_date):
                for shift_type_list in schedule[prev_date].values():
                    if shift_type_list is not None and staff_name in shift_type_list:
                        return True
        return False

    def assign_no_reply_shifts(self, schedule, all_dates, no_reply_members):
        workdays = [date for date in all_dates if not self.is_weekend(date)]

        for member in no_reply_members:
            shifts_needed = 2
            random.shuffle(workdays)

            for date in workdays:
                if shifts_needed <= 0:
                    break

                valid_shifts = (
                    ["opening", "middle"]
                    if self.is_monday(date)
                    else ["opening", "middle", "closing"]
                )
                random.shuffle(valid_shifts)

                for shift in valid_shifts:
                    if schedule[date][shift] is not None and len(
                        schedule[date][shift]
                    ) < self.get_staff_requirement(date, shift):
                        schedule[date][shift].append(member)
                        shifts_needed -= 1
                        break

            if shifts_needed == 2:
                for date in workdays:
                    valid_shifts = (
                        ["opening", "middle"]
                        if self.is_monday(date)
                        else ["opening", "middle", "closing"]
                    )
                    for shift in valid_shifts:
                        if schedule[date][shift] is not None and len(
                            schedule[date][shift]
                        ) < self.get_staff_requirement(date, shift):
                            schedule[date][shift].append(member)
                            break
                    else:
                        continue
                    break

    def validate_schedule(self, schedule, all_dates):
        for date in all_dates:
            if self.is_monday(date):
                schedule[date]["closing"] = None

                for shift in ["opening", "middle"]:
                    if schedule[date][shift] is not None:
                        required = self.get_staff_requirement(date, shift)
                        schedule[date][shift] = schedule[date][shift][:required]
            else:
                for shift_type in ["opening", "middle", "closing"]:
                    if schedule[date][shift_type] is not None:
                        required = self.get_staff_requirement(date, shift_type)
                        schedule[date][shift_type] = schedule[date][shift_type][
                            :required
                        ]

        return schedule

    def format_date_str(self, day):
        return f"{day}. {self.MONTH_NAME[:3].lower()}"

    def apply_excel_formatting(self, ws, all_dates, all_members):
        ws.freeze_panes = "B2"

        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        header_fill = PatternFill(
            start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"
        )
        header_font = Font(bold=True, size=11)

        names_fill = PatternFill(
            start_color="F5F5F5", end_color="F5F5F5", fill_type="solid"
        )
        names_font = Font(bold=True, size=11)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in range(2, len(all_members) + 2):
            cell = ws.cell(row=row, column=1)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

            name = cell.value
            if name not in self.no_reply_members:
                cell.font = names_font
                cell.fill = names_fill

        last_col = len(all_dates) + 3
        for row in range(2, len(all_members) + 2):
            for col in range(2, last_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")

        sum_row = len(all_members) + 3
        total_rows = [sum_row, sum_row + 1, sum_row + 2, sum_row + 3]

        for row in total_rows:
            cell = ws.cell(row=row, column=1)
            cell.font = Font(bold=True)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.fill = header_fill

            for col in range(2, last_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.column_dimensions["A"].width = 30
        for col in range(2, ws.max_column + 1):
            if ws.cell(row=1, column=col).value == "WEEKEND":
                ws.column_dimensions[get_column_letter(col)].width = 15
            else:
                ws.column_dimensions[get_column_letter(col)].width = 12

        for row in range(1, len(all_members) + 6):
            ws.row_dimensions[row].height = 22

    def _assign_first_shifts(self, schedule, dates, staff_availability, all_members):
        assigned_members = set()

        for staff_name, availability in staff_availability.items():
            if staff_name in assigned_members:
                continue

            random.shuffle(availability)
            for date, shifts in availability:
                if not self.is_weekend(date) and not self.check_consecutive_days(
                    schedule, staff_name, date, dates
                ):
                    valid_shifts = self.get_available_shifts(date)
                    random.shuffle(valid_shifts)

                    for shift in valid_shifts:
                        if schedule[date][shift] is not None and len(
                            schedule[date][shift]
                        ) < self.get_staff_requirement(date, shift):
                            schedule[date][shift].append(staff_name)
                            assigned_members.add(staff_name)
                            break
                    if staff_name in assigned_members:
                        break

        workdays = [date for date in dates if not self.is_weekend(date)]
        for member in self.no_reply_members:
            if member in assigned_members:
                continue

            random.shuffle(workdays)
            for date in workdays:
                valid_shifts = (
                    ["opening", "middle"]
                    if self.is_monday(date)
                    else ["opening", "middle", "closing"]
                )
                random.shuffle(valid_shifts)

                for shift in valid_shifts:
                    if schedule[date][shift] is not None and len(
                        schedule[date][shift]
                    ) < self.get_staff_requirement(date, shift):
                        schedule[date][shift].append(member)
                        assigned_members.add(member)
                        break
                if member in assigned_members:
                    break

    def _assign_second_shifts(self, schedule, dates, staff_availability, all_members):
        unfilled_slots = 0
        for date in dates:
            if self.is_weekend(date):
                continue
            for shift_type in ["opening", "middle", "closing"]:
                if schedule[date][shift_type] is not None:
                    required = self.get_staff_requirement(date, shift_type)
                    current = len(schedule[date][shift_type])
                    unfilled_slots += max(0, required - current)

        if unfilled_slots > 0:
            for staff_name, availability in staff_availability.items():
                if self._count_shifts(schedule, staff_name) >= 1:
                    random.shuffle(availability)
                    for date, shifts in availability:
                        if not self.is_weekend(
                            date
                        ) and not self.check_consecutive_days(
                            schedule, staff_name, date, dates
                        ):
                            valid_shifts = self.get_available_shifts(date)
                            random.shuffle(valid_shifts)

                            for shift in valid_shifts:
                                if schedule[date][shift] is not None and len(
                                    schedule[date][shift]
                                ) < self.get_staff_requirement(date, shift):
                                    schedule[date][shift].append(staff_name)
                                    break
                            if self._count_shifts(schedule, staff_name) >= 2:
                                break

            workdays = [date for date in dates if not self.is_weekend(date)]
            for member in self.no_reply_members:
                if self._count_shifts(schedule, member) >= 1:
                    random.shuffle(workdays)
                    for date in workdays:
                        valid_shifts = (
                            ["opening", "middle"]
                            if self.is_monday(date)
                            else ["opening", "middle", "closing"]
                        )
                        random.shuffle(valid_shifts)

                        for shift in valid_shifts:
                            if schedule[date][shift] is not None and len(
                                schedule[date][shift]
                            ) < self.get_staff_requirement(date, shift):
                                schedule[date][shift].append(member)
                                break
                        if self._count_shifts(schedule, member) >= 2:
                            break

    def _count_shifts(self, schedule, staff_name):
        return sum(
            1
            for date in schedule
            for shift_type, staff_list in schedule[date].items()
            if staff_list is not None and staff_name in staff_list
        )

    def create_schedule(self):
        with open(self.USERPATH + "members.txt", "r") as f:
            all_members = [line.strip() for line in f if line.strip()]
        df = pd.read_csv(self.FILEPATH)

        date_cols = [
            col
            for col in df.columns
            if f"{self.MONTH_NAME[:3].lower()} -" in col.lower()
        ]
        work_dates = [col.split("[")[-1].split("]")[0].strip() for col in date_cols]

        all_dates = []
        current_date = None
        for i, date in enumerate(work_dates):
            if current_date:
                current_day = int(current_date.split(".")[0])
                next_day = int(date.split(".")[0])
                current = datetime(self.YEAR, self.MONTH, current_day)
                while current.weekday() < 5 and current.day < next_day:
                    current = current + timedelta(days=1)
                    if current.day < next_day:
                        weekend_date = f"{current.day}. {self.MONTH_NAME[:3].lower()}"
                        all_dates.append(weekend_date)
            all_dates.append(date)
            current_date = date

        schedule = {}
        for date in all_dates:
            if self.is_weekend(date):
                schedule[date] = {"opening": None, "middle": None, "closing": None}
            else:
                schedule[date] = {
                    "opening": [],
                    "middle": [],
                    "closing": None if self.is_monday(date) else [],
                }

        staff_availability = {}
        responding_members = set()

        for _, row in df.iterrows():
            input_name = row["Navn og etternavn"]
            matched_name = self.find_member_match(input_name, all_members)
            if matched_name in all_members:
                responding_members.add(matched_name)
                availability = []
                for date, col in zip(work_dates, date_cols):
                    shifts = self.parse_shifts(row[col], date)
                    if shifts:
                        availability.append((date, shifts))
                staff_availability[matched_name] = availability

        self.no_reply_members = set(all_members) - responding_members

        self._assign_first_shifts(schedule, work_dates, staff_availability, all_members)

        self._assign_second_shifts(
            schedule, work_dates, staff_availability, all_members
        )

        schedule = self.validate_schedule(schedule, all_dates)

        wb = Workbook()
        ws = wb.active
        ws.title = "Schedule"

        legend_row = 1
        ws.cell(row=legend_row, column=len(all_dates) + 5, value="Shift Colors:")
        for idx, (shift_name, info) in enumerate(self.shifts.items(), 1):
            cell = ws.cell(
                row=legend_row + idx,
                column=len(all_dates) + 5,
                value=f"{shift_name.capitalize()} ({info['time']})",
            )
            cell.fill = PatternFill(
                start_color=info["color"], end_color=info["color"], fill_type="solid"
            )

        ws["A1"] = "Name"
        for idx, date in enumerate(all_dates, 2):
            cell = ws.cell(
                row=1, column=idx, value="WEEKEND" if self.is_weekend(date) else date
            )
            if self.is_weekend(date):
                cell.fill = PatternFill(
                    start_color=self.weekend_color,
                    end_color=self.weekend_color,
                    fill_type="solid",
                )

        for row_idx, name in enumerate(all_members, 2):
            cell = ws.cell(row=row_idx, column=1, value=name)

            if name in self.no_reply_members:
                cell.fill = PatternFill(
                    start_color=self.no_reply_color,
                    end_color=self.no_reply_color,
                    fill_type="solid",
                )
                cell.font = Font(color="FFFFFF")
            elif any(review["possible_match"] == name for review in self.manual_review):
                cell.fill = PatternFill(
                    start_color="D3D3D3", end_color="D3D3D3", fill_type="solid"
                )

            for col_idx, date in enumerate(all_dates, 2):
                cell = ws.cell(row=row_idx, column=col_idx)
                if self.is_weekend(date):
                    cell.fill = PatternFill(
                        start_color=self.weekend_color,
                        end_color=self.weekend_color,
                        fill_type="solid",
                    )
                else:
                    if self.is_monday(date):
                        for shift, staff_list in schedule[date].items():
                            if (
                                shift != "closing"
                                and staff_list is not None
                                and name in staff_list
                            ):
                                cell.fill = PatternFill(
                                    start_color=self.shifts[shift]["color"],
                                    end_color=self.shifts[shift]["color"],
                                    fill_type="solid",
                                )
                    else:
                        for shift, staff_list in schedule[date].items():
                            if staff_list is not None and name in staff_list:
                                cell.fill = PatternFill(
                                    start_color=self.shifts[shift]["color"],
                                    end_color=self.shifts[shift]["color"],
                                    fill_type="solid",
                                )

        for date in all_dates:
            if self.is_monday(date):
                col_idx = all_dates.index(date) + 2
                for row_idx in range(2, len(all_members) + 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.fill.start_color.index == "C6FFB4":
                        cell.fill = PatternFill(fill_type=None)

        staff_shifts = {member: 0 for member in all_members}
        for date in all_dates:
            for shift_type in ["opening", "middle", "closing"]:
                if schedule[date].get(shift_type) is not None:
                    for staff in schedule[date][shift_type]:
                        if staff in staff_shifts:
                            staff_shifts[staff] += 1

        ws.cell(row=1, column=len(all_dates) + 2, value="Total Shifts")
        ws.cell(row=1, column=len(all_dates) + 3, value="Available Days")
        for row_idx, name in enumerate(all_members, 2):
            ws.cell(row=row_idx, column=len(all_dates) + 2, value=staff_shifts[name])
            available_count = len(staff_availability.get(name, []))
            ws.cell(row=row_idx, column=len(all_dates) + 3, value=available_count)

        if self.manual_review:
            ws_review = wb.create_sheet("Manual Review")
            ws_review["A1"] = "Input Name"
            ws_review["B1"] = "Possible Match"
            ws_review["C1"] = "Confidence"
            ws_review["D1"] = "Available Dates"

            for idx, review in enumerate(self.manual_review, 2):
                ws_review[f"A{idx}"] = review["input_name"]
                ws_review[f"B{idx}"] = review["possible_match"]
                ws_review[f"C{idx}"] = review.get("confidence", "N/A")

                if review["input_name"] in self.unmatched_availability:
                    avail_text = []
                    for date, shifts in self.unmatched_availability[
                        review["input_name"]
                    ].items():
                        if shifts:
                            avail_text.append(f"{date}: {', '.join(shifts)}")
                    ws_review[f"D{idx}"] = "\n".join(avail_text)

        if self.manual_review:
            ws_review = wb.create_sheet("Manual Review")
            ws_review["A1"] = "Input Name"
            ws_review["B1"] = "Possible Match"
            ws_review["C1"] = "Confidence"
            ws_review["D1"] = "Available Dates"

            for idx, review in enumerate(self.manual_review, 2):
                ws_review[f"A{idx}"] = review["input_name"]
                ws_review[f"B{idx}"] = review["possible_match"]
                ws_review[f"C{idx}"] = review.get("confidence", "N/A")

                if review["input_name"] in self.unmatched_availability:
                    avail_text = []
                    for date, shifts in self.unmatched_availability[
                        review["input_name"]
                    ].items():
                        if shifts:
                            avail_text.append(f"{date}: {', '.join(shifts)}")
                    ws_review[f"D{idx}"] = "\n".join(avail_text)

        sum_row = len(all_members) + 3
        ws.cell(row=sum_row, column=1, value="SUM OF SHIFTS")

        for col_idx, date in enumerate(all_dates, 2):
            if not self.is_weekend(date):
                opening_count = (
                    len(schedule[date]["opening"])
                    if schedule[date]["opening"] is not None
                    else 0
                )
                middle_count = (
                    len(schedule[date]["middle"])
                    if schedule[date]["middle"] is not None
                    else 0
                )
                closing_count = (
                    len(schedule[date]["closing"])
                    if schedule[date]["closing"] is not None
                    else 0
                )

                ws.cell(
                    row=sum_row,
                    column=col_idx,
                    value=opening_count + middle_count + closing_count,
                )

                ws.cell(row=sum_row + 1, column=1, value="OPENING")
                ws.cell(row=sum_row + 1, column=col_idx, value=opening_count)

                ws.cell(row=sum_row + 2, column=1, value="MIDDAY")
                ws.cell(row=sum_row + 2, column=col_idx, value=middle_count)

                ws.cell(row=sum_row + 3, column=1, value="CLOSING")
                ws.cell(row=sum_row + 3, column=col_idx, value=closing_count)

        self.apply_excel_formatting(ws, all_dates, all_members)

        save_path = (
            f"{self.USERPATH}{self.MONTH_NAME.lower()}_schedule_{self.YEAR}.xlsx"
        )
        print(f"Saving schedule to: {save_path}")
        wb.save(save_path)


def main():
    scheduler = BarScheduler()
    scheduler.create_schedule()


if __name__ == "__main__":
    main()
