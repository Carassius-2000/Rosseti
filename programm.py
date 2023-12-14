import tkinter
from datetime import datetime
from tkinter import filedialog, messagebox

import customtkinter
import joblib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from customtkinter import (
    CTk,
    CTkButton,
    CTkFont,
    CTkFrame,
    CTkLabel,
    CTkOptionMenu,
    IntVar,
    StringVar
)
from pymongo import MongoClient

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")


class Application(CTk):
    """
    Main application window.

    Attributes
    ----------
    filetypes : tuple(tuple[str, str])
        File types for opening and saving files.
    forecast_horizons : list[str]
        List of forecast horizons.
    data : pandas.DataFrame
        Temp data storage.
    """

    __filetypes = (("Книга Excel", "*.xlsx"),)

    __forecast_horizons = [
        "На один день",
        "На два дня",
        "На три дня",
        "На четыре дня",
        "На пять дней",
        "На шесть дней",
        "На семь дней"
    ]

    def __init__(self):
        """
        Initializes main application window.

        Sets up main application window by creating and configuring
        various UI elements.
        """
        super().__init__()
        font = CTkFont(size=16, weight="bold")
        self.protocol("WM_DELETE_WINDOW", self.__close_app)
        self.title("Прогнозирование потребления электроэнергии")
        self.resizable(0, 0)
        self.geometry("1240x250")
        self.toplevel_window = None

        self.__data = None

        self.grid_columnconfigure((1, 2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        load_frame = CTkFrame(self, corner_radius=10)
        load_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

        load_label = CTkLabel(
            load_frame, text="Выберите способ загрузки данных:", font=font
        )
        load_label.grid(row=0, column=0, padx=20, pady=20)

        load_menu_variable = StringVar(value="Загрузить из БД")  # значение по умолчанию
        self.__load_combobox = CTkOptionMenu(
            load_frame,
            values=["Загрузить из Excel", "Загрузить из БД"],
            variable=load_menu_variable,
            width=300,
            height=32,
            font=font
        )
        self.__load_combobox.grid(row=1, column=0, padx=20, pady=20)

        load_data_button = CTkButton(
            load_frame,
            width=300,
            height=32,
            text="Загрузить данные",
            command=self.__get_data,
            font=font
        )
        load_data_button.grid(row=2, column=0, padx=20, pady=20)

        prediction_frame = CTkFrame(self, corner_radius=10)
        prediction_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        prediction_label = CTkLabel(
            prediction_frame,
            text="На сколько дней вперёд нужно делать прогноз:",
            font=font
        )
        prediction_label.grid(row=0, column=1, padx=20, pady=20)

        days_menu_variable = StringVar(
            value=self.__forecast_horizons[0]
        )  # значение по умолчанию
        self.__days_combobox = CTkOptionMenu(
            prediction_frame,
            values=self.__forecast_horizons,
            variable=days_menu_variable,
            width=300,
            state="disabled",
            height=32,
            font=font
        )
        self.__days_combobox.grid(row=1, column=1, padx=20, pady=20)

        self.__get_predictions_button = CTkButton(
            prediction_frame,
            width=300,
            height=32,
            text="Получить прогнозы",
            state="disabled",
            command=self.__get_predictions,
            font=font
        )
        self.__get_predictions_button.grid(row=2, column=1, padx=20, pady=20)

        save_frame = CTkFrame(self, corner_radius=10)
        save_frame.grid(row=0, column=2, sticky="nsew", padx=20, pady=20)

        self.__visualization_button = CTkButton(
            save_frame,
            width=300,
            height=32,
            text="Построить график",
            state="disabled",
            command=self.__visualization,
            font=font
        )
        self.__visualization_button.grid(row=0, column=2, padx=20, pady=20)

        self.__save_to_db_button = CTkButton(
            save_frame,
            width=300,
            height=32,
            text="Сохранить в БД",
            state="disabled",
            command=self.__save_to_db,
            font=font
        )
        self.__save_to_db_button.grid(row=1, column=2, padx=20, pady=20)

        self.__save_to_excel_button = CTkButton(
            save_frame,
            width=300,
            height=32,
            text="Сохранить в Excel",
            state="disabled",
            command=self.__save_to_excel,
            font=font
        )
        self.__save_to_excel_button.grid(row=2, column=2, padx=20, pady=20)

    def __get_data(self) -> None:
        """
        Retrieves data based on selection made in load_combobox.
        """
        if self.__load_combobox.get() == "Загрузить из Excel":
            self.__data = self.__load_from_excel()
        elif self.__load_combobox.get() == "Загрузить из БД":
            self.__data = self.__load_from_db()
        if self.__data is not None:
            self.__days_combobox.configure(state=tkinter.NORMAL)
            self.__get_predictions_button.configure(state=tkinter.NORMAL)
            messagebox.showinfo("Информация", "Данные успешно загружены.")

    @classmethod
    def __load_from_excel(cls) -> pd.DataFrame:
        """
        Loads data from Excel.

        Returns
        ----------
        `~pandas.DataFrame`
        """
        file_path: str = cls.__fetch_file_path()
        if not file_path:
            messagebox.showerror(" ", "Вы не выбрали Excel файл")
            return None
        data: pd.DataFrame = pd.read_excel(file_path, index_col=0)
        data.index = data.index.round("H")
        data = data.iloc[-24 * 7 :]
        return data

    @classmethod
    def __fetch_file_path(cls) -> str:
        """
        Opens dialog to select Excel file and returns file path.

        Returns
        -------
        `~str`
        """
        file_path: str = filedialog.askopenfilename(
            title="Открыть файл",
            initialdir="/",
            filetypes=cls.__filetypes,
            defaultextension=".xlsx"
        )
        return file_path

    @staticmethod
    def __load_from_db() -> pd.DataFrame:
        """
        Get data from MongoDB

        Returns
        ----------
        `~pandas.DataFrame`: Загруженные данные из MongoDB.
        """
        with MongoClient(serverSelectionTimeoutMS=1000) as client:
            db = client["rosseti"]
            collection = db["electricity_consumption"]
            data = collection.find(limit=24 * 7, sort=[("timestamp", -1)])

        data = pd.DataFrame(data).sort_values(by=["timestamp"])
        data.rename(
            columns={
                "timestamp": "Дата и время",
                "electricity_consumption": "Электропотребление"
            },
            inplace=True,
        )
        data.index = data["Дата и время"].dt.round("H")
        data.drop(columns=["_id", "Дата и время"], inplace=True)
        return data

    def __get_predictions(self) -> None:
        """
        Get predictions for next day's electricity consumption.
        """
        forecast_horizon: int = (
            self.__forecast_horizons.index(self.__days_combobox.get()) + 1
        )
        last_available_day: pd.Series = self.__data.index[-1]
        forecast_day_begin: pd.Timestamp = last_available_day + pd.DateOffset(days=1)
        prediction_range: pd.DatetimeIndex = pd.date_range(
            start=forecast_day_begin, periods=24 * forecast_horizon, freq="h"
        )
        mask = self.__create_mask_fill_na(
            self.__data["Электропотребление"].iloc[-24:], forecast_horizon
        )
        prediction_data: pd.DataFrame = pd.DataFrame(
            {"Электропотребление": mask.values}, index=prediction_range
        )
        prediction_data.index.name = "Дата и время"
        X: pd.DataFrame = pd.concat([self.__data, prediction_data])
        X = self.__preprocessing_data(X)
        model = joblib.load("regression.model")
        prediction_data["Электропотребление"] = np.round(model.predict(X), 3)

        self.__visualization_button.configure(state=tkinter.NORMAL)
        self.__save_to_db_button.configure(state=tkinter.NORMAL)
        self.__save_to_excel_button.configure(state=tkinter.NORMAL)
        self.__data = prediction_data
        messagebox.showinfo("Информация", "Прогнозы успешно получены.")

    @staticmethod
    def __create_mask_fill_na(last_day: pd.DataFrame, num_days: int) -> pd.Series:
        """
        Creates a mask to fill missing values in a DataFrame with data from last available day.

        Parameters
        ----------
        last_day : pandas.DataFrame
            Input client data.
        num_days : int
            Number of days to fill missing values.

        Returns:
        ----------
        `~pandas.Series`
        """
        return pd.concat([last_day for _ in range(num_days)], ignore_index=True)

    @staticmethod
    def __create_times_of_day(data: pd.DataFrame) -> pd.Series:
        """
        Creates a column for time of day.
        Rows are grouped by hours.
        - 0 (night) [0 - 5]
        - 1 (morning) [6 - 11]
        - 2 (lunch) [12 - 17]
        - 3 (evening) [18 - 23]

        Parameters
        ----------
        data : pandas.DataFrame
            Input data.

        Returns
        ----------
        `~pandas.Series`
        """
        return pd.cut(data.index.hour, bins=4, labels=range(4))

    @classmethod
    def __add_time_features(cls, data: pd.DataFrame) -> None:
        """
        Adds time-related features to input data.

        - Hour [0 - 23]
        - Time of Day [0 - 3]
        - Day of the week [0 - 6]
        - Weekend [0 - 1]
        - Month [1 - 12]
        - Day of the year [1 - 366]

        Parameters
        ----------
        data : pandas.DataFrame
            Input data.
        """
        data["Час"] = data.index.hour.astype("category")
        data["Период времени суток"] = cls.__create_times_of_day(data).astype(
            "category"
        )
        data["День недели"] = data.index.dayofweek.astype("category")
        data["Выходной"] = (
            data["День недели"].isin([5, 6]).astype("int").astype("category")
        )
        data["Месяц"] = data.index.month.astype("category")
        data["День в году"] = data.index.dayofyear

    @staticmethod
    def __add_lag_features(data: pd.DataFrame) -> None:
        """
        Creates lag features related to target variable.

        - Electricity consumption lag 1 day (time shift by 24 hours)
        - Electricity consumption lag 7 days (time shift by 168 hours)

        Parameters
        ----------
        data : pandas.DataFrame
            Input data.
        """
        data["Электропотребление лаг 1 день"] = data["Электропотребление"].shift(24)
        data["Электропотребление лаг 7 дней"] = data["Электропотребление"].shift(24 * 7)

    @classmethod
    def __preprocessing_data(cls, data: pd.DataFrame) -> pd.DataFrame:
        """
        Preprocesses given data by rounding "Date/Time" column to nearest hour,
        adding time features, adding lag features, dropping rows with missing values, and returning
        preprocessed data.

        Parameters
        ----------
        data : pandas.DataFrame
            Input data to be preprocessed.

        Returns
        ----------
        `~pandas.DataFrame`
        """
        cls.__add_time_features(data)
        cls.__add_lag_features(data)
        data.dropna(inplace=True)
        data.drop(columns=["Электропотребление"], inplace=True)
        return data

    def __visualization(self) -> None:
        """
        Visualizes data using a plot.
        """
        _, ax = plt.subplots(figsize=(12, 6))
        ax.plot(self.__data["Электропотребление"], marker="o")
        ax.set_title(
            f"Прогноз {self.__days_combobox.get().lower()} вперёд", fontsize=18
        )
        ax.tick_params(axis="both", labelsize=18)
        ax.set_xlabel("Дата и время", fontsize=18)
        ax.set_ylabel("Потребление электроэнергии (МВт * ч)", fontsize=18)
        ax.grid(axis="y")
        plt.show()

    def __save_to_db(self) -> None:
        """
        Save the data to a MongoDB database.
        """
        data = self.__data.copy().reset_index().to_dict("records")
        result = [
            {
                "metadata": {"report_date": datetime.now()},
                "electricity_consumption": element["Электропотребление"],
                "timestamp": element["Дата и время"]
            }
            for element in data
        ]

        with MongoClient(serverSelectionTimeoutMS=1000) as client:
            db = client["rosseti"]
            collection = db["reports"]
            collection.insert_many(result)

        messagebox.showinfo(" ", "Прогнозы успешно записаны в БД")

    def __save_to_excel(self) -> None:
        """
        Save data to an Excel file.
        Opens a file dialog to allow the user to choose a file name and location for saving data.
        """
        filename: str = filedialog.asksaveasfilename(
            title="Сохранить файл",
            initialdir="/",
            filetypes=self.__filetypes,
            defaultextension=".xlsx"
        )
        self.__data.to_excel(filename, sheet_name="Лист1")
        messagebox.showinfo(" ", f"Прогнозы успешно записаны в {filename}")

    def __close_app(self) -> None:
        """
        Closes application.

        Checks if user wants to exit application.
        If user confirms, application is destroyed.
        """
        if messagebox.askyesno("Выход из приложения", "Хотите выйти из приложения?"):
            self.destroy()


if __name__ == "__main__":
    root = Application()
    root.eval("tk::PlaceWindow . center")
    root.mainloop()
