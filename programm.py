import tkinter
from datetime import datetime
from tkinter import filedialog, messagebox
import customtkinter
import joblib
import matplotlib.pyplot as plt
import pandas as pd
from customtkinter import (
    CTk,
    CTkButton,
    CTkFont,
    CTkFrame,
    CTkLabel,
    CTkOptionMenu,
    StringVar
)
from pymongo import MongoClient
from pymongo.errors import ServerSelectionTimeoutError


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")


class Application(CTk):
    __filetypes: tuple[tuple[str, str]] = (("Книга Excel", "*.xlsx"),)
    __forecast_horizons: list[str] = [
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
        self.wm_iconbitmap("prediction.ico")
        self.toplevel_window = None

        self.__data = None
        self.__mongo_client = None

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
    def __load_from_excel(cls):
        """
        Loads data from Excel.

        Returns
        ----------
        `~pandas.DataFrame`
        """
        file_path: str = cls.__fetch_file_path("load")
        if not file_path:
            messagebox.showerror("Ошибка", "Вы не выбрали Excel файл")
            return
        data = pd.read_excel(file_path, index_col=0)
        data = DataProcessor().postprocess_data_from_excel(data)
        return data

    @classmethod
    def __fetch_file_path(cls, option: str) -> str:
        """
        Opens dialog to select Excel file and returns file path.

        Returns
        -------
        `~str`
        """
        if option == "load":
            file_path: str = filedialog.askopenfilename(
                title="Открыть файл",
                filetypes=cls.__filetypes,
                defaultextension=".xlsx"
            )
        elif option == "save":
            file_path: str = filedialog.asksaveasfilename(
                title="Сохранить файл",
                filetypes=cls.__filetypes,
                defaultextension=".xlsx"
            )
        return file_path

    def __load_from_db(self):
        """
        Get data from MongoDB

        Returns
        ----------
        `~pandas.DataFrame`
        """
        self.__mongo_client = MongoDBDriver()
        data = self.__mongo_client.load_data(
            db_name="rosseti", collection_name="electricity_consumption"
        )
        try:
            data = pd.DataFrame(data)
        except ServerSelectionTimeoutError:
            messagebox.showerror(
                "Ошибка",
                "Не удалось загрузить данные из БД.\nПроверьте подключение к серверу."
            )
            return
        data = DataProcessor().postprocess_data_from_db(data)
        return data

    def __get_predictions(self) -> None:
        """
        Get predictions for next day's electricity consumption.
        """
        forecast_horizon: int = (
            self.__forecast_horizons.index(self.__days_combobox.get()) + 1
        )
        future_dataframe = DataProcessor().make_future_dataframe(
            forecast_horizon, self.__data
        )
        X = pd.concat([self.__data, future_dataframe])
        X = DataProcessor().preprocessing_data(X)
        model = joblib.load("regression.model")
        predictions = model.predict(X).round(3)
        future_dataframe["Электропотребление"] = predictions

        self.__visualization_button.configure(state=tkinter.NORMAL)
        self.__save_to_db_button.configure(state=tkinter.NORMAL)
        self.__save_to_excel_button.configure(state=tkinter.NORMAL)
        self.__data = future_dataframe
        messagebox.showinfo("Информация", "Прогнозы успешно получены.")

    def __visualization(self) -> None:
        """
        Visualizes data using a plot.
        """
        drawer = Drawer(self.__data)
        forecast_horizon: int = self.__days_combobox.get().lower()
        drawer.line_plot(horizon_size=forecast_horizon)

    def __save_to_db(self) -> None:
        """
        Save the data to a MongoDB database.
        """
        data_to_save = DataProcessor().prepare_data_for_saving(self.__data.copy())
        if self.__mongo_client is None:
            self.__mongo_client = MongoDBDriver() 
        self.__mongo_client.save_data(
            data=data_to_save, db_name="rosseti", collection_name="reports"
        )

    def __save_to_excel(self) -> None:
        """
        Save data to an Excel file.
        Opens a file dialog to allow the user to choose a file name and location for saving data.
        """
        file_path: str = self.__fetch_file_path("save")
        if not file_path:
            messagebox.showerror(
                "Ошибка", "Вы не выбрали путь для сохранения Excel файла"
            )
            return
        self.__data.to_excel(file_path, sheet_name="Лист1")
        messagebox.showinfo("Информация", f"Прогнозы успешно записаны в {file_path}")

    def __close_app(self) -> None:
        """
        Closes application.

        Checks if user wants to exit application.
        If user confirms, application is destroyed.
        """
        if messagebox.askyesno("Выход из приложения", "Хотите выйти из приложения?"):
            self.destroy()


class Drawer:
    def __init__(self, data):
        """
        Initialize the Drawer class.

        Parameters
        ----------
        data :
            Data for plotting.
        """
        self.__data = data

    def line_plot(
        self,
        horizon_size: str,
        plot_size: tuple[int, int] = (12, 6),
        font_size: int = 18
    ) -> None:
        """
        Visualizes data using a line plot.

        Parameters
        ----------
        horizon_size : str
        plot_size : tuple[int, int]
        font_size : int
        """
        _, ax = plt.subplots(figsize=plot_size)
        ax.plot(self.__data, marker="o")
        ax.set_title(f"Прогноз {horizon_size} вперёд", fontsize=font_size)
        ax.tick_params(axis="both", labelsize=font_size)
        ax.set_xlabel("Дата и время", fontsize=font_size)
        ax.set_ylabel("Потребление электроэнергии (МВт * ч)", fontsize=font_size)
        ax.grid(axis="y")
        plt.show()


class MongoDBDriver:
    def __init__(self):
        """Initialize the MongoDBDriver class."""
        self.__connection = MongoClient(
            serverSelectionTimeoutMS=1_000, maxPoolSize=None, waitQueueTimeoutMS=1_000
        )

    def load_data(self, db_name: str, collection_name: str):
        """
        Load data from database collection.

        Parameters
        ----------
        db_name : str
            Name of database.
        collection_name: str
            Name of collection within database.

        Returns
        ----------
        `~pymongo.cursor.Cursor`
        """
        db = self.__connection[db_name]
        collection = db[collection_name]
        data = collection.find(limit=24 * 7, sort=[("timestamp", -1)])
        return data

    def save_data(self, db_name: str, collection_name: str, data) -> None:
        """
        Save data to database and collection using the provided DataFrame.

        Parameters
        ----------
        db_name : str
            Name of database to save data to.
        collection_name: str
            Name of collection within database to save data to.
        data : pd.DataFrame
            DataFrame containing data to be saved.
        """
        db = self.__connection[db_name]
        collection = db[collection_name]
        try:
            collection.insert_many(data)
        except ServerSelectionTimeoutError:
            messagebox.showerror(
                "Ошибка",
                "Не удалось записать данные в БД.\nПроверьте подключение к серверу."
            )
        else:
            messagebox.showinfo("Информация", "Прогнозы успешно записаны в БД")


class DataProcessor:
    @staticmethod
    def postprocess_data_from_excel(data):
        """
        Postprocess data from Excel file.

        Parameters
        ----------
        data : pandas.DataFrame
            Input client data.

        Returns
        ----------
        `~pandas.DataFrame`
        """
        data.index = data.index.round("H")
        data = data.iloc[-24 * 7 :]
        return data

    @staticmethod
    def postprocess_data_from_db(data):
        """
        Postprocesses data retrieved from database.

        Parameters
        ----------
        data : data retrieved from database.

        Returns
        ----------
        `~pandas.DataFrame`
        """
        data = pd.DataFrame(data).sort_values(by=["timestamp"])
        data.rename(
            columns={
                "timestamp": "Дата и время",
                "electricity_consumption": "Электропотребление"
            },
            inplace=True
        )
        data.index = data["Дата и время"].dt.round("H")
        data.drop(columns=["_id", "Дата и время"], inplace=True)
        return data

    @staticmethod
    def __create_mask_fill_na(last_day, num_days: int):
        """
        Creates a mask to fill missing values in a DataFrame with data from last available day.

        Parameters
        ----------
        last_day : pandas.DataFrame
            Input client data.
        num_days : int
            Number of days to fill missing values.

        Returns
        ----------
        `~pandas.Series`
        """
        return pd.concat([last_day for _ in range(num_days)], ignore_index=True)

    @classmethod
    def make_future_dataframe(cls, forecast_horizon: int, data):
        """
        Generate a future dataframe for making predictions.

        Parameters
        ----------
            forecast_horizon : int
                Number of days to forecast into the future.
            data : pd.DataFrame
                Input data for generating future dataframe.

        Returns
        ----------
        `~pandas.DataFrame`
        """
        last_available_day = data.index[-1]
        forecast_day_begin = last_available_day + pd.DateOffset(hours=1)
        prediction_range = pd.date_range(
            start=forecast_day_begin, periods=24 * forecast_horizon, freq="h"
        )
        last_day_by_hour = data["Электропотребление"].iloc[-24:]
        mask = cls.__create_mask_fill_na(last_day_by_hour, forecast_horizon)
        prediction_data = pd.DataFrame(
            {"Электропотребление": mask.values}, index=prediction_range
        )
        prediction_data.index.name = "Дата и время"
        return prediction_data

    @staticmethod
    def __create_times_of_day(data):
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
    def __add_time_features(cls, data) -> None:
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
    def __add_lag_features(data) -> None:
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
    def preprocessing_data(cls, data):
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

    @staticmethod
    def prepare_data_for_saving(data):
        """
        Prepare input data for saving to MongoDB.

        Parameters
        ----------
            data: input data to be prepared for saving.

        Returns
        ----------
        `~pandas.DataFrame`
        """
        data_to_save = data.reset_index().to_dict("records")
        data_to_save = [
            {
                "metadata": {"report_date": datetime.now()},
                "electricity_consumption": element["Электропотребление"],
                "timestamp": element["Дата и время"]
            }
            for element in data_to_save
        ]
        return data_to_save


if __name__ == "__main__":
    root = Application()
    root.eval("tk::PlaceWindow . center")
    root.mainloop()
