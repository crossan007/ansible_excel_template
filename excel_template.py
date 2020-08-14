# Copyright: (c) 2020, Charles crossan
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

from __future__ import (absolute_import, division, print_function)
__metaclass__ = type

import os
import glob
import shutil
import stat
import tempfile
import re
import zipfile

from ansible import constants as C
from ansible.config.manager import ensure_type
from ansible.errors import AnsibleError, AnsibleFileNotFound, AnsibleAction, AnsibleActionFail
from ansible.module_utils._text import to_bytes, to_text, to_native
from ansible.module_utils.parsing.convert_bool import boolean
from ansible.module_utils.six import string_types
from ansible.plugins.action import ActionBase


class ActionModule(ActionBase):

    def run(self, tmp=None, task_vars=None):
        ''' handler for replace operations '''

        if task_vars is None:
            task_vars = dict()

        action_plugin_result = super(ActionModule, self).run(tmp, task_vars)
        del tmp  # tmp no longer has any effect

        # Options type validation
        # stings
        for s_type in ('src', 'dest'):
            if s_type in self._task.args:
                value = ensure_type(self._task.args[s_type], 'string')
                if value is not None and not isinstance(value, string_types):
                    raise AnsibleActionFail("%s is expected to be a string, but got %s instead" % (s_type, type(value)))
                self._task.args[s_type] = value

       # assign to local vars for ease of use
       
        src = self._task.args.get('src')
        dest = self._task.args.get('dest')

        print("Templating from " + src)

        res_args = dict()

        local_tempdir = tempfile.mkdtemp(dir=C.DEFAULT_LOCAL_TMP)

        # first; unzip the xlsx workbook:
        with zipfile.ZipFile(src, 'r') as zip_ref:
            zip_ref.extractall(local_tempdir) 

        #print("Extracted to " + local_tempdir)

        sheets_dir = os.path.join(local_tempdir,"xl","worksheets","*.xml")

        #print("Scaning " + sheets_dir)

        common_args = {}
        files_to_template = [os.path.join(local_tempdir,"xl","sharedStrings.xml")] #glob.glob(sheets_dir)


        for file_to_template in files_to_template:
            new_task = self._task.copy()
            new_task.args['mode'] = self._task.args.get('mode', None)
            new_task.args.update(
                    dict(
                        src=file_to_template,
                        dest=file_to_template
                    ),
                )
            template_action = self._shared_loader_obj.action_loader.get('template',
                                                                        task=new_task,
                                                                        connection=self._connection,
                                                                        play_context=self._play_context,
                                                                        loader=self._loader,
                                                                        templar=self._templar,
                                                                        shared_loader_obj=self._shared_loader_obj)
            template_result = template_action.run(task_vars=task_vars)
            print(template_result)


        zf = zipfile.ZipFile(dest, "w")
        for dirname, subdirs, files in os.walk(local_tempdir):
            zf.write(dirname)
            for filename in files:
                file_path = os.path.join(dirname, filename)
                rel_path = file_path.replace(local_tempdir,"")
                zf.write(file_path,rel_path)
        zf.close()

        action_plugin_result["changed"] = False
        
        return action_plugin_result