# HOW TO USE THIS MODULE
# OPEN INM AND STUDY
# - Create an INMStudy() object and store it in a variable (e.g. inmstudy1)
# - Open a new INM window with the open_inm() method
# - Open study with open_study() and pass it the study directory
# you wish to open
#
# SET GRIDS
# - Open the grid menu with open_grid_setup()
# - Create new grids by calling the set_grid() method any number of times
#
# SET RUN OPTIONS
# - Set run options with set_run_options()
#
# RUN STUDY
# - Run study with run_study()
#
# CLOSE STUDY AND INM
# - Close study with close_study()
# - Close INM with close_inm()

import os
import sys
import time

import pywinauto
import pywinauto.timings
from pywinauto import handleprops
from pywinauto.application import Application
from pywinauto.findwindows import find_windows

__author__ = 'Thomas Vandenhede'


class INMAuto:
    """
    This class provides a set of methods to interact with the INM GUI.
    """

    def __init__(self, inm_exe_path):
        self.app = Application()
        self.inm_exe_path = inm_exe_path
        self.wtitle = 'INM.*'  # typical title: 'INM 7.0'
        self.main_window = None
        self.noise_metric = None
        self.path_to_study = None
        self.study_folder = None

    @staticmethod
    def __path_to_dir(path):
        """
        This function transforms a path into directories and formats
        the directories in preparation for use in 'Open' and
        'Export As' Menus.

        """
        path = path.lower()
        directory = path.split('\\')
        directory[0] += "\\"
        return directory

    def open_inm(self):
        """
        This function simply opens INM.

        """
        # close any INM window if already open
        os.system("taskkill /f /im inm.exe 2> nul")
        try:
            # open INM
            self.app.start(self.inm_exe_path)
            self.main_window = self.app.window_(title_re=self.wtitle)
            self.main_window.Wait('ready')
        except (pywinauto.application.AppStartError,
                pywinauto.findbestmatch.MatchError) as err:
            print(err)
            print('Error: %s' % sys.exc_info()[0].__name__)
            print(sys.exc_info()[1])
            exit(1)

    def open_study(self, path_to_study):
        """
        This function opens the study specified as path_to_study_.

        """
        try:
            # Close any open study
            self.close_study()

            # split and format path in preparation for use in inm open menu
            self.path_to_study = path_to_study
            self.study_folder = os.path.basename(path_to_study)
            dir_path = self.__path_to_dir(path_to_study)

            # Open study specified in path_to_study
            self.main_window.MenuItem('File->Open Study...').Click()
            self.main_window.WaitNot('ready')
            w_open = self.app.top_window_()

            # browse through directories and accept
            w_open['ListBox'].SetFocus()
            for d in dir_path:
                item_texts = w_open['ListBox'].ItemTexts()
                item_index = [name.lower() for name in item_texts].index(d)
                w_open['ListBox'].Select(item_index)
                w_open.TypeKeys('{ENTER}')

            w_open['OKButton'].Click()

            # update INM window title with study folder
            self.main_window = self.app['- [Study %s]' % path_to_study.upper()]
        except (pywinauto.application.AppStartError,
                pywinauto.findbestmatch.MatchError) as err:
            print(err)
            self.close_all_windows()
            self.open_study(path_to_study)

    def open_grid_setup(self):
        """
        This function simply opens the 'Grid Setup' menu in INM.

        """
        try:
            # open 'Grid Options' selection window
            self.close_all_windows()
            self.main_window.MenuItem('Run->Grid Setup...').Click()

            # select the topmost scenario in the 'Scenario Select' window
            w_select = self.app['Scenario Select']
            w_select['ListBox'].Select(0)
            w_select['OKButton'].Click()

            self.rearrange_windows_in_cascade()

            # delete all grids in the 'Grid Points Setup' window
            w_grid_setup = self.app.top_window_()
            w_grid_setup['ListBox'].SetFocus()
            item_count = w_grid_setup['ListBox'].ItemCount()
            for i in range(item_count):
                self.click_menu_item('Edit->Delete Records')
        except (pywinauto.application.AppStartError,
                pywinauto.findbestmatch.MatchError) as err:
            print(err)
            self.close_all_windows()
            self.open_grid_setup()

    def set_grid(self, grid_setup):
        """
        This function configures a grid in the 'Grid Setup' Menu.

        """
        try:
            # create new grid and set grid type
            w_grid_setup = self.app.top_window_()
            self.click_menu_item('Edit->Add Record')
            w_grid_setup['Grip TypeComboBox'].Select(grid_setup.grid_type)

            # set threshold and 'do percent of time' option
            if grid_setup.relative_threshold:
                w_grid_setup['Relative ThresholdRadioButton'].Click()
                w_grid_setup['Ambient + Delta (dB)Edit'].SetEditText(
                    grid_setup.relative_threshold)
            else:
                w_grid_setup['Fixed Threshold (dB)RadioButton'].Click()
                w_grid_setup['Fixed Threshold (dB)Edit'].SetEditText(
                    grid_setup.fixed_threshold)

            if grid_setup.do_percent_of_time:
                checkbox = w_grid_setup['Do Percent of Time (hr)CheckBox']
                if checkbox.GetCheckState() == 0:
                    w_grid_setup['Do Percent of Time (hr)CheckBox'].Click()
                w_grid_setup['Do Percent of Time (hr)Edit'].SetEditText(
                    grid_setup.do_percent_of_time)
            else:
                w_grid_setup['Do Percent of Time (hr)CheckBox'].UnCheck()

            # set x, y, i, j [or lat, long (not implemented!)]
            if (grid_setup.grid_type == 'Contour' or
                        grid_setup.grid_type == 'Standard' or
                        grid_setup.grid_type == 'Detailed'):
                if grid_setup.coordinates == 'X/Y':
                    w_grid_setup['X/YRadioButton'].Click()
                    w_grid_setup['X (nmi)Edit'].SetEditText(grid_setup.x)
                    w_grid_setup['Y (nmi)Edit'].SetEditText(grid_setup.y)
                    w_grid_setup['I (nmi)Edit'].SetEditText(grid_setup.i)
                    w_grid_setup['J (nmi)Edit'].SetEditText(grid_setup.j)
                else:
                    pass

                w_grid_setup['Grid Rotation Angle (deg)Edit'].SetEditText(
                    grid_setup.grid_rotation_angle)

            if (grid_setup.grid_type == 'Standard' or
                        grid_setup.grid_type == 'Detailed'):
                w_grid_setup['Grid IdEdit'].SetEditText(grid_setup.grid_id)
                w_grid_setup['IEdit'].SetEditText(grid_setup.nb_pts_i)
                w_grid_setup['JEdit'].SetEditText(grid_setup.nb_pts_j)
        except (pywinauto.application.AppStartError,
                pywinauto.findbestmatch.MatchError) as err:
            print(err)
            self.close_all_windows()
            self.set_grid(grid_setup)

    def set_run_options(self, run_options):
        """
        This function sets the run options in the 'Run Options' menu

        """
        try:
            # Open 'Run Options' Menu and close other windows
            self.close_all_windows()
            self.click_menu_item('Run->Run Options')
            self.rearrange_windows_in_cascade()
            w_runoptions = self.app.top_window_()

            # Select desired parameters
            w_runoptions['Run TypeComboBox'].Select(run_options.run_type)
            if run_options.do_terrain:
                w_runoptions['Do TerrainCheckBox'].Check()
            w_runoptions['Lateral AttenuationComboBox'].Select(
                run_options.lateral_attenuation)
            if run_options.use_bank_angle:
                w_runoptions['Use Bank AngleCheckBox'].Check()

            # format noise metric name to match case and length in INM
            self.noise_metric = run_options.noise_metric.ljust(6).upper()

            if run_options.run_type == 'Single-Metric':
                w_runoptions['Noise MetricComboBox'].Select(self.noise_metric)

                if run_options.do_contours:
                    # Contour
                    w_runoptions['Do ContoursCheckBox'].Check()
                    if run_options.use_boundary_file:
                        w_runoptions['Use Boundary FileCheckBox'].Check()
                    if run_options.fixed_grid:
                        w_runoptions['Fixed GridRadioButton'].Click()
                        if run_options.fixed_spacing:
                            w_runoptions['Fixed SpacingRadioButton'].Click()
                            w_runoptions['SpacingEdit'].SetEditText(
                                run_options.spacing)
                        else:
                            w_runoptions['RefinementRadioButton'].Click()
                            w_runoptions['RefinementComboBox'].Select(
                                run_options.refinement)
                    else:
                        w_runoptions['Recursive Grid'].Click()
                        w_runoptions['RefinementComboBox'].Select(
                            run_options.refinement)
                        w_runoptions['ToleranceEdit'].SetEditText(
                            run_options.tolerance)
                        w_runoptions['Low CutoffEdit'].SetEditText(
                            run_options.low_cutoff)
                        w_runoptions['High CutoffEdit'].SetEditText(
                            run_options.high_cutoff)

                # Pop and Loc Points
                if run_options.do_population_points:
                    w_runoptions['Do Population PointsCheckBox'].Check()
                if run_options.do_location_points:
                    w_runoptions['Do Location PointsCheckBox'].Check()

                # Grid
                if run_options.do_standard_grids:
                    w_runoptions['Do Standard GridsCheckBox'].Check()
                if run_options.do_detailed_grids:
                    w_runoptions['Do Detailed GridsCheckBox'].Check()
                    if run_options.save_all_flights:
                        w_runoptions['Save 100% FlightsCheckBox'].Check()

                # Calculate Metrics
                if run_options.dnl:
                    w_runoptions['DNLCheckBox'].Check()
                if run_options.cnel:
                    w_runoptions['CNELCheckBox'].Check()
                if run_options.laeq:
                    w_runoptions['LAEQCheckBox'].Check()
                if run_options.laeqd:
                    w_runoptions['LAEQDCheckBox'].Check()
                if run_options.laeqn:
                    w_runoptions['LAEQNCheckBox'].Check()
                if run_options.sel:
                    w_runoptions['SELCheckBox'].Check()
                if run_options.lamax:
                    w_runoptions['LAMAXCheckBox'].Check()
                if run_options.tala:
                    w_runoptions['TALACheckBox'].Check()
                if run_options.nef:
                    w_runoptions['NEFCheckBox'].Check()
                if run_options.wecpnl:
                    w_runoptions['WECPNLCheckBox'].Check()
                if run_options.epnl:
                    w_runoptions['EPNLCheckBox'].Check()
                if run_options.pnltm:
                    w_runoptions['PNLTMCheckBox'].Check()
                if run_options.tapnl:
                    w_runoptions['TAPNLCheckBox'].Check()
                if run_options.cexp:
                    w_runoptions['CEXPCheckBox'].Check()
                if run_options.lcmax:
                    w_runoptions['LCMAXCheckBox'].Check()
                if run_options.talc:
                    w_runoptions['TALCCheckBox'].Check()
        except (pywinauto.application.AppStartError,
                pywinauto.findbestmatch.MatchError) as err:
            print(err)
            self.close_all_windows()
            self.set_run_options(run_options)

    def run_study(self):
        """
        Simply runs the INM study.

        """
        try:
            # Run Start... from 'Run' Menu
            self.close_all_windows()
            self.main_window.MenuItem('Run->Run Start...').Click()
            self.main_window.WaitNot('ready')
            w_runstart = self.app['Run Start']

            # Select first scenario in list and proceed...
            w_runstart['Scenario ListListBox'].Select(0)
            w_runstart['Include --- >ListBox'].Click()
            w_runstart['OKButton'].Click()

            # Close all intermediary windows until the 'Run Status' window
            # shows up
            try:
                while True:
                    self.main_window.Maximize()
                    self.main_window.WaitNot('ready')
                    dlg = self.app.top_window_()
                    # Close the 'Warning' window that shows up if study has
                    # already been run
                    if dlg.WindowText() == 'Warning':
                        dlg['OuiButton'].Click()
                    # Close any information window
                    elif dlg.WindowText() == 'INM 7.0':
                        print(dlg['Static2'].Texts()[0])
                        dlg['OKButton'].Click()
                    # 'Run status' dialog appears... (note: this dialog's
                    # title is an empty string, ie '')
                    else:
                        print('Running study...')
                        break
            except OverflowError as err:
                print(err)
            # Wait until the 'Run Status' window closes
            self.main_window.Wait('ready', timeout=300, retry_interval=0.01)
            print('Run finished!')
        except (pywinauto.application.AppStartError,
                pywinauto.findbestmatch.MatchError) as err:
            print(err)
            self.close_all_windows()
            self.run_study()

    # 'Run Options' Menu
    def export_output(self, export_options):
        """
        This function exports all the output specified in its arguments.

        """
        try:
            self.get_noise_metric_from_run_options_menu()
            self.set_output_noise_metric()

            # Start Outputting...
            if export_options.output_graphics:
                self.export_output_graphics()

            if export_options.contour_points:
                self.export_contour_points(export_options.file_type)

            if export_options.contour_area_and_pop:
                self.export_contour_area_and_pop(export_options.file_type)

            if export_options.area_contour_coverage:
                self.export_area_contour_coverage(export_options.file_type)

            if export_options.standard_grids:
                self.export_standard_grids(export_options.file_type)

            if export_options.detailed_grids:
                self.export_detailed_grids(export_options.file_type)

            if export_options.noise_at_pop_points:
                self.export_noise_at_pop_point(export_options.file_type)

            if export_options.noise_at_loc_points:
                self.export_noise_at_loc_point(export_options.file_type)

            if export_options.scenario_run_input_report:
                pass  # TODO: implement the export function for input report

            if export_options.flight_path_report:
                self.export_flight_path_report()
        except (pywinauto.application.AppStartError,
                pywinauto.findbestmatch.MatchError) as err:
            print(err)
            self.close_all_windows()
            self.export_output(export_options)

    def get_noise_metric_from_run_options_menu(self):
        self.click_menu_item('Run->Run Options')
        dlg = self.app.top_window_()
        noise_metric = dlg['Noise MetricComboBox'].Texts()[0]
        self.noise_metric = noise_metric.ljust(6).upper()
        self.close_all_windows()

    def set_output_noise_metric(self):
        # set output metric to be the same as 'Run Options' noise metric
        self.click_menu_item('Output->Output Setup')
        w_outputsetup = self.app.top_window_()
        w_outputsetup['MetricComboBox'].Select(self.noise_metric)
        self.close_all_windows()

    def export_output_graphics(self):
        self.main_window.MenuItem('Output->Output Graphics...').Click()
        dlg_title = 'Output Select'
        self.__select(dlg_title)

        while True:
            w_top = self.app.top_window_()
            if w_top['OKButton'].Exists(0.005, 0.001):
                w_top['OKButton'].Click()
            else:
                if w_top.ChildWindow(
                        title_re='Output',
                        class_name='AfxFrameOrView42'
                ).Exists(0.005, 0.001):
                    break

        self.rearrange_windows_in_cascade()
        self.__export_graphics()

    def export_contour_points(self, file_type):
        self.main_window.MenuItem('Output->Contour Points...').Click()
        dlg_title = 'Output Select'
        self.__select(dlg_title)

        while True:
            w_top = self.app.top_window_()
            if w_top['OKButton'].Exists(0.005, 0.001):
                w_top['OKButton'].Click()
            else:
                if w_top.ChildWindow(
                        title_re='Contour Points',
                        class_name='AfxFrameOrView42'
                ).Exists(0.005, 0.001):
                    break

        self.rearrange_windows_in_cascade()
        self.__export(file_type)

    def export_contour_area_and_pop(self, file_type):
        self.main_window.MenuItem('Output->Contour Area and Pop...').Click()
        dlg_title = 'Output Select'
        self.__select(dlg_title)

        # wait until 'Contour Area and Population' window appears
        while True:
            w_top = self.app.top_window_()
            if w_top['OKButton'].Exists(0.005, 0.001):
                w_top['OKButton'].Click()
            else:
                if w_top.ChildWindow(
                        title_re='Contour Area and Population',
                        class_name='AfxFrameOrView42').Exists(0.005,
                                                              0.001):
                    break

        self.rearrange_windows_in_cascade()
        self.__export(file_type)

    def export_area_contour_coverage(self, file_type):
        self.main_window.MenuItem(
            'Output->Area Contour Coverage...').Click()
        if self.app.window_(title='ERROR').Exists():
            self.app.top_window_().OKButton.Click()

        self.rearrange_windows_in_cascade()
        self.__export(file_type)

    def export_standard_grids(self, file_type):
        self.click_menu_item('Output->Standard Grids...')
        dlg_title = 'Scenario Select'
        self.__select(dlg_title)

        self.rearrange_windows_in_cascade()
        self.__export(file_type)

    def export_detailed_grids(self, file_type):
        self.click_menu_item('Output->Detailed Grids...')
        dlg_title = 'Scenario Select'
        self.__select(dlg_title)

        self.rearrange_windows_in_cascade()
        self.__export(file_type)

    def export_noise_at_pop_point(self, file_type):
        self.click_menu_item('Output->Noise at Pop Points...')
        dlg_title = 'Scenario Select'
        self.__select(dlg_title)

        self.rearrange_windows_in_cascade()
        self.__export(file_type)

    def export_noise_at_loc_point(self, file_type):
        # open 'Noise at Loc Points...' menu
        self.click_menu_item('Output->Noise at Loc Points...')
        dlg_title = 'Scenario Select'
        self.__select(dlg_title)

        self.rearrange_windows_in_cascade()
        self.__export(file_type)

    def export_flight_path_report(self):
        self.click_menu_item('Output->Flight Path Report...')
        dlg_title = 'Scenario Select'
        self.__select(dlg_title)

        # Click OK button in dialog that appears to confirm
        self.app.top_window_()['OKButton'].Click()

    def __select(self, w_title):
        w_select = self.app[w_title]
        w_select['ListBox'].Select(0)
        w_select['OKButton'].Click()

    def __export(self, file_type):
        # create new folder in 'OUTPUT1' with same name as noise metric
        output_dir = os.path.join(
            self.path_to_study, 'OUTPUT1', self.noise_metric.strip())
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Open 'Export As...' window
        self.click_menu_item('File->Export As...')
        w_export = self.app.window_(title_re='Export As*')

        # add study folder to name of output file
        edit_text = w_export['File NameEdit'].Texts()[0]
        export_file = "%s_%s" % (edit_text, self.study_folder)
        w_export['List Files or TypeComboBox'].Select(file_type)
        w_export['File NameEdit'].SetEditText(export_file)

        # Get output directory (exclude the drive from the split path)
        dir_path = self.__path_to_dir(output_dir)
        if w_export['ListBox'].GetItemFocus() == 0:
            dir_path.pop(0)

        # browse through directories and accept
        w_export['ListBox'].Click()
        print(dir_path)
        for d in dir_path:
            item_texts = w_export['ListBox'].ItemTexts()
            item_index = [name.lower() for name in item_texts].index(d)
            print(d, item_index)
            w_export['ListBox'].Select(item_index)
            w_export.TypeKeys('{ENTER}')

        # if confirm dialog appears chose to override existing file
        if len(find_windows(title_re='Export As.*')) != 0:
            self.app.top_window_()['ReplaceButton'].Click()
        self.close_all_windows()
        print("%s output created for %s"
              % (export_file, self.noise_metric.strip()))

    def __export_graphics(self):
        # create new folder in 'OUTPUT1' with same name as noise metric
        output_dir = os.path.join(
            self.path_to_study, 'OUTPUT1', self.noise_metric.strip())
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

            # Open 'Export As...' window
            self.click_menu_item('File->Export as ShapeFile...')
        w_export = self.app.window_(title_re='Export As Shapefile')

        # add study folder to name of output file
        unit = 'feet'
        w_export['Export UnitsComboBox'].Select(unit)
        w_export['BrowseButton'].Click()

        # 'Directories' window appears
        w_directories = self.app.window_(title_re='Directories')
        w_directories['ListBox'].SetFocus()

        # browse through directories and accept
        dir_ = self.__path_to_dir(output_dir)
        if w_directories['ListBox'].GetItemFocus() == 0:
            dir_.pop(0)
        w_directories.TypeKeys('{HOME}')
        for d in dir_:
            w_directories['ListBox'].Select(d)
            w_directories.TypeKeys('{ENTER}')
        w_directories['OKButton'].Click()

        w_export['OKButton'].Click()

    def close_study(self):
        """
        Closes the study.
        """
        item_string = 'File->Close Study'
        menu_text = self.main_window.MenuItem(item_string).Text()
        if 'Close' in menu_text:
            self.click_menu_item(item_string)
        self.main_window = self.app['INM.*']

    def close_inm(self):
        """
        Closes INM.
        """
        self.click_menu_item('File->Exit')

    def rearrange_windows_in_cascade(self):
        """
        Self explanatory. Rearranges all windows of the GUI in cascade.
        """
        self.click_menu_item('Window->Cascade')

    def close_all_windows(self):
        self.click_menu_item('Window->Close All')

    def click_menu_item(self, item_string):
        """
        Close all windows in the GUI.
        """
        menu_item = self.main_window.MenuItem(item_string)
        if menu_item.IsEnabled():
            menu_item.Click()


class StudyOptions(object):
    def __init__(self):
        self._params = dict()
        self.reset()

    def reset(self):
        raise NotImplementedError()


class GridSetup(StudyOptions):
    """
    The grid_setup class holds all the information about a grid setup. There
    can be multiple grid_setup instances for one INM study. The
    purpose of this class is to be instantiated in the INMStudy class.

    """
    def __init__(self):
        StudyOptions.__init__(self)

    def reset(self):
        """
        Resets all parameters to their default values.

        :return:
        """
        self._params['grid_type'] = 'Location'
        self._params['grid_id'] = None
        self._params['coordinates'] = 'X/Y'
        self._params['x'] = -8.0
        self._params['y'] = -8.0
        self._params['i'] = 16.0
        self._params['j'] = 16.0
        self._params['nb_pts_i'] = 2
        self._params['nb_pts_j'] = 2
        self._params['grid_rotation_angle'] = 0.0
        self._params['fixed_threshold'] = 85.0
        self._params['relative_threshold'] = None
        self._params['do_percent_of_time'] = None

    def get_grid_setup_dict(self):
        return self._params

    @property
    def grid_type(self):
        return self._params['grid_type']

    @grid_type.setter
    def grid_type(self, grid_type):
        self._params['grid_type'] = grid_type

    @property
    def grid_id(self):
        return self._params['grid_id']

    @grid_id.setter
    def grid_id(self, grid_id):
        self._params['grid_id'] = grid_id

    @property
    def coordinates(self):
        return self._params['coordinates']

    @coordinates.setter
    def coordinates(self, coordinates):
        self._params['coordinates'] = coordinates

    @property
    def x(self):
        return self._params['x']

    @x.setter
    def x(self, x):
        self._params['x'] = x

    @property
    def y(self):
        return self._params['y']

    @y.setter
    def y(self, y):
        self._params['y'] = y

    @property
    def i(self):
        return self._params['i']

    @i.setter
    def i(self, i):
        self._params['i'] = i

    @property
    def j(self):
        return self._params['j']

    @j.setter
    def j(self, j):
        self._params['j'] = j

    @property
    def nb_pts_i(self):
        return self._params['nb_pts_i']

    @nb_pts_i.setter
    def nb_pts_i(self, nb_pts_i):
        self._params['nb_pts_i'] = nb_pts_i

    @property
    def nb_pts_j(self):
        return self._params['nb_pts_j']

    @nb_pts_j.setter
    def nb_pts_j(self, nb_pts_j):
        self._params['nb_pts_j'] = nb_pts_j

    @property
    def grid_rotation_angle(self):
        return self._params['grid_rotation_angle']

    @grid_rotation_angle.setter
    def grid_rotation_angle(self, grid_rotation_angle):
        self._params['grid_rotation_angle'] = grid_rotation_angle

    @property
    def fixed_threshold(self):
        return self._params['fixed_threshold']

    @fixed_threshold.setter
    def fixed_threshold(self, fixed_threshold):
        self._params['fixed_threshold'] = fixed_threshold

    @property
    def relative_threshold(self):
        return self._params['relative_threshold']

    @relative_threshold.setter
    def relative_threshold(self, relative_threshold):
        self._params['relative_threshold'] = relative_threshold

    @property
    def do_percent_of_time(self):
        return self._params['do_percent_of_time']

    @do_percent_of_time.setter
    def do_percent_of_time(self, do_percent_of_time):
        self._params['do_percent_of_time'] = do_percent_of_time


class RunOptions(StudyOptions):
    """
    The RunOptions class holds all the information about a grid setup. The
    purpose of this class is to be instantiated in the INMStudy class.

    """
    def __init__(self):
        StudyOptions.__init__(self)

    def reset(self):
        """
        Resets all parameters to their default values.

        :return:
        """
        self._params['run_type'] = 'Single-Metric'
        self._params['noise_metric'] = 'LAMAX '
        self._params['do_terrain'] = False
        self._params['lateral_attenuation'] = 'All-Soft-Ground'
        self._params['use_bank_angle'] = False
        self._params['do_contours'] = False
        self._params['use_boundary_file'] = False
        self._params['refinement'] = 4
        self._params['low_cutoff'] = 55.0
        self._params['tolerance'] = 0.25
        self._params['high_cutoff'] = 85.0
        self._params['fixed_grid'] = True
        self._params['fixed_spacing'] = True
        self._params['spacing'] = 1000.0
        self._params['do_population_points'] = False
        self._params['do_location_points'] = False
        self._params['do_standard_grids'] = False
        self._params['do_detailed_grids'] = False
        self._params['save_all_flights'] = False
        self._params['dnl'] = False
        self._params['cnel'] = False
        self._params['laeq'] = False
        self._params['laeqd'] = False
        self._params['laeqn'] = False
        self._params['sel'] = False
        self._params['lamax'] = False
        self._params['tala'] = False
        self._params['nef'] = False
        self._params['wecpnl'] = False
        self._params['epnl'] = False
        self._params['pnltm'] = False
        self._params['tapnl'] = False
        self._params['cexp'] = False
        self._params['lcmax'] = False
        self._params['talc'] = False

    def get_run_options_dict(self):
        return self._params

    @property
    def run_type(self):
        return self._params['run_type']

    @run_type.setter
    def run_type(self, run_type):
        self._params['run_type'] = run_type

    @property
    def noise_metric(self):
        return self._params['noise_metric']

    @noise_metric.setter
    def noise_metric(self, noise_metric):
        self._params['noise_metric'] = noise_metric

    @property
    def do_terrain(self):
        return self._params['do_terrain']

    @do_terrain.setter
    def do_terrain(self, do_terrain):
        self._params['do_terrain'] = do_terrain

    @property
    def lateral_attenuation(self):
        return self._params['lateral_attenuation']

    @lateral_attenuation.setter
    def lateral_attenuation(self, lateral_attenuation):
        self._params['lateral_attenuation'] = lateral_attenuation

    @property
    def use_bank_angle(self):
        return self._params['use_bank_angle']

    @use_bank_angle.setter
    def use_bank_angle(self, use_bank_angle):
        self._params['use_bank_angle'] = use_bank_angle

    @property
    def do_contours(self):
        return self._params['do_contours']

    @do_contours.setter
    def do_contours(self, do_contours):
        self._params['do_contours'] = do_contours

    @property
    def use_boundary_file(self):
        return self._params['use_boundary_file']

    @use_boundary_file.setter
    def use_boundary_file(self, use_boundary_file):
        self._params['use_boundary_file'] = use_boundary_file

    @property
    def refinement(self):
        return self._params['refinement']

    @refinement.setter
    def refinement(self, refinement):
        self._params['refinement'] = refinement

    @property
    def low_cutoff(self):
        return self._params['low_cutoff']

    @low_cutoff.setter
    def low_cutoff(self, low_cutoff):
        self._params['low_cutoff'] = low_cutoff

    @property
    def tolerance(self):
        return self._params['tolerance']

    @tolerance.setter
    def tolerance(self, tolerance):
        self._params['tolerance'] = tolerance

    @property
    def high_cutoff(self):
        return self._params['high_cutoff']

    @high_cutoff.setter
    def high_cutoff(self, high_cutoff):
        self._params['high_cutoff'] = high_cutoff

    @property
    def fixed_grid(self):
        return self._params['fixed_grid']

    @fixed_grid.setter
    def fixed_grid(self, fixed_grid):
        self._params['fixed_grid'] = fixed_grid

    @property
    def fixed_spacing(self):
        return self._params['fixed_spacing']

    @fixed_spacing.setter
    def fixed_spacing(self, fixed_spacing):
        self._params['fixed_spacing'] = fixed_spacing

    @property
    def spacing(self):
        return self._params['spacing']

    @spacing.setter
    def spacing(self, spacing):
        self._params['spacing'] = spacing

    @property
    def do_population_points(self):
        return self._params['do_population_points']

    @do_population_points.setter
    def do_population_points(self, do_population_points):
        self._params['do_population_points'] = do_population_points

    @property
    def do_location_points(self):
        return self._params['do_location_points']

    @do_location_points.setter
    def do_location_points(self, do_location_points):
        self._params['do_location_points'] = do_location_points

    @property
    def do_standard_grids(self):
        return self._params['do_standard_grids']

    @do_standard_grids.setter
    def do_standard_grids(self, do_standard_grids):
        self._params['do_standard_grids'] = do_standard_grids

    @property
    def do_detailed_grids(self):
        return self._params['do_detailed_grids']

    @do_detailed_grids.setter
    def do_detailed_grids(self, do_detailed_grids):
        self._params['do_detailed_grids'] = do_detailed_grids

    @property
    def save_all_flights(self):
        return self._params['save_all_flights']

    @save_all_flights.setter
    def save_all_flights(self, save_all_flights):
        self._params['save_all_flights'] = save_all_flights

    @property
    def dnl(self):
        return self._params['dnl']

    @dnl.setter
    def dnl(self, dnl):
        self._params['dnl'] = dnl

    @property
    def cnel(self):
        return self._params['cnel']

    @cnel.setter
    def cnel(self, cnel):
        self._params['cnel'] = cnel

    @property
    def laeq(self):
        return self._params['laeq']

    @laeq.setter
    def laeq(self, laeq):
        self._params['laeq'] = laeq

    @property
    def laeqd(self):
        return self._params['laeqd']

    @laeqd.setter
    def laeqd(self, laeqd):
        self._params['laeqd'] = laeqd

    @property
    def laeqn(self):
        return self._params['laeqn']

    @laeqn.setter
    def laeqn(self, laeqn):
        self._params['laeqn'] = laeqn

    @property
    def sel(self):
        return self._params['sel']

    @sel.setter
    def sel(self, sel):
        self._params['sel'] = sel

    @property
    def lamax(self):
        return self._params['lamax']

    @lamax.setter
    def lamax(self, lamax):
        self._params['lamax'] = lamax

    @property
    def tala(self):
        return self._params['tala']

    @tala.setter
    def tala(self, tala):
        self._params['tala'] = tala

    @property
    def nef(self):
        return self._params['nef']

    @nef.setter
    def nef(self, nef):
        self._params['nef'] = nef

    @property
    def wecpnl(self):
        return self._params['wecpnl']

    @wecpnl.setter
    def wecpnl(self, wecpnl):
        self._params['wecpnl'] = wecpnl

    @property
    def epnl(self):
        return self._params['epnl']

    @epnl.setter
    def epnl(self, epnl):
        self._params['epnl'] = epnl

    @property
    def pnltm(self):
        return self._params['pnltm']

    @pnltm.setter
    def pnltm(self, pnltm):
        self._params['pnltm'] = pnltm

    @property
    def tapnl(self):
        return self._params['tapnl']

    @tapnl.setter
    def tapnl(self, tapnl):
        self._params['tapnl'] = tapnl

    @property
    def cexp(self):
        return self._params['cexp']

    @cexp.setter
    def cexp(self, cexp):
        self._params['cexp'] = cexp

    @property
    def lcmax(self):
        return self._params['lcmax']

    @lcmax.setter
    def lcmax(self, lcmax):
        self._params['lcmax'] = lcmax

    @property
    def talc(self):
        return self._params['talc']

    @talc.setter
    def talc(self, talc):
        self._params['talc'] = talc


class ExportOptions(StudyOptions):
    """
    The ExportOptions class holds all the information about a grid setup. The
    purpose of this class is to be instantiated in the INMStudy class.

    """
    def __init__(self):
        StudyOptions.__init__(self)

    def reset(self):
        """
        Resets all parameters to their default values.

        :return:
        """
        self._params['output_graphics'] = None
        self._params['contour_points'] = None
        self._params['contour_area_and_pop'] = None
        self._params['area_contour_coverage'] = None
        self._params['standard_grids'] = None
        self._params['detailed_grids'] = None
        self._params['noise_at_pop_points'] = None
        self._params['noise_at_loc_points'] = None
        self._params['scenario_run_input_report'] = None
        self._params['flight_path_report'] = None
        self._params['file_type'] = None

    def get_export_options_dict(self):
        return self._params

    @property
    def output_graphics(self):
        return self._params['output_graphics']

    @output_graphics.setter
    def output_graphics(self, output_graphics):
        self._params['output_graphics'] = output_graphics

    @property
    def contour_points(self):
        return self._params['contour_points']

    @contour_points.setter
    def contour_points(self, contour_points):
        self._params['contour_points'] = contour_points

    @property
    def contour_area_and_pop(self):
        return self._params['contour_area_and_pop']

    @contour_area_and_pop.setter
    def contour_area_and_pop(self, contour_area_and_pop):
        self._params['contour_area_and_pop'] = contour_area_and_pop

    @property
    def area_contour_coverage(self):
        return self._params['area_contour_coverage']

    @area_contour_coverage.setter
    def area_contour_coverage(self, area_contour_coverage):
        self._params['area_contour_coverage'] = area_contour_coverage

    @property
    def standard_grids(self):
        return self._params['standard_grids']

    @standard_grids.setter
    def standard_grids(self, standard_grids):
        self._params['standard_grids'] = standard_grids

    @property
    def detailed_grids(self):
        return self._params['detailed_grids']

    @detailed_grids.setter
    def detailed_grids(self, detailed_grids):
        self._params['detailed_grids'] = detailed_grids

    @property
    def noise_at_pop_points(self):
        return self._params['noise_at_pop_points']

    @noise_at_pop_points.setter
    def noise_at_pop_points(self, noise_at_pop_points):
        self._params['noise_at_pop_points'] = noise_at_pop_points

    @property
    def noise_at_loc_points(self):
        return self._params['noise_at_loc_points']

    @noise_at_loc_points.setter
    def noise_at_loc_points(self, noise_at_loc_points):
        self._params['noise_at_loc_points'] = noise_at_loc_points

    @property
    def scenario_run_input_report(self):
        return self._params['scenario_run_input_report']

    @scenario_run_input_report.setter
    def scenario_run_input_report(self, scenario_run_input_report):
        self._params['scenario_run_input_report'] = scenario_run_input_report

    @property
    def flight_path_report(self):
        return self._params['flight_path_report']

    @flight_path_report.setter
    def flight_path_report(self, flight_path_report):
        self._params['flight_path_report'] = flight_path_report

    @property
    def file_type(self):
        return self._params['file_type']

    @file_type.setter
    def file_type(self, file_type):
        self._params['file_type'] = file_type


class INMStudy(object):
    """
    This class is a wrapper for the lower level INMAuto class. It calls an
    INMAuto instance's methods in the proper order in order to perform a
    complete INM study (open study, create grids, set run options,
    run study, export output, close study).

    """

    def __init__(self, inm_exe_path, study_path):
        self.inm = INMAuto(inm_exe_path)
        self.study_folder = study_path

        # path must be absolute
        self.path_to_study = os.path.join(
            os.getcwd(), 'INM Studies', self.study_folder)

    def run_scenario(self, grids, run_options, export_options=None):
        self.inm.open_inm()
        self.inm.open_study(self.path_to_study)

        # Setup grids
        self.inm.open_grid_setup()
        for grid in grids:
            self.inm.set_grid(grid)

        # Set 'Run Options' and run studies
        self.inm.set_run_options(run_options)
        self.inm.run_study()

        # Export only if export options specified
        if export_options:
            self.inm.export_output(export_options)

        # Close study
        self.inm.close_study()

    def __is_inm_open(self):
        pass

    def __is_study_open(self):
        pass
