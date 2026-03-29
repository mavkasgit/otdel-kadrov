    def _draw_calendar(self) -> None:
        self._update_widget_bootstyle()
        self._set_title()
        self._current_month_days()
        self.frm_dates = ttk.Frame(self.frm_calendar)
        self.frm_dates.pack(fill=BOTH, expand=YES)

        for row, weekday_list in enumerate(self.monthdays):
            for col, day in enumerate(weekday_list):
                self.frm_dates.columnconfigure(col, weight=1)
                if day == 0:
                    ttk.Label(
                        master=self.frm_dates,
                        text=self.monthdates[row][col].day,
                        anchor=CENTER,
                        padding=5,
                        bootstyle=SECONDARY,
                    ).grid(row=row, column=col, sticky=NSEW)
                else:
                    if all(
                            [
                                day == self.date_selected.day,
                                self.date.month == self.date_selected.month,
                                self.date.year == self.date_selected.year,
                            ]
                    ):
                        day_style = "secondary-toolbutton"
                    else:
                        day_style = f"{self.bootstyle}-calendar"

                    def selected(x=row, y=col):
                        self._on_date_selected(x, y)

                    btn = ttk.Radiobutton(
                        master=self.frm_dates,
                        variable=self.datevar,
                        value=day,
                        text=day,
                        takefocus=True,
                        bootstyle=day_style,
                        padding=5,
                        command=selected,
                    )
                    btn.grid(row=row, column=col, sticky=NSEW)

