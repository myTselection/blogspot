
/**
 * CustodianDailyRollingFileAppender.java
 * Adapted from the Apache Log4j DailyRollingFileAppender to extend the functionality
 * of the existing class so that the user can limit the number of log backups
 * and compress the backups to conserve disk space.
 * @author Ryan Kimber, minor changes by myTselection.blogspot.com

 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.myt.common.logging;

import org.apache.log4j.FileAppender;
import org.apache.log4j.Layout;
import org.apache.log4j.helpers.LogLog;
import org.apache.log4j.spi.LoggingEvent;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import java.net.MalformedURLException;
import java.net.URL;

import java.text.SimpleDateFormat;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Locale;
import java.util.TimeZone;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


/**
 * CustodianDailyRollingFileAppender is based on {@link org.apache.log4j.appender.DailyRollingFileAppender} so most of
 * the configuration options can be taken from the documentation on that class.
 */
public class CustodianDailyRollingFileAppender extends FileAppender {
    private static final TimeZone GMT = TimeZone.getTimeZone("GMT");

    // The code assumes that the following constants are in a increasing sequence.
    private static final int TOP_OF_TROUBLE = -1;
    private static final int TOP_OF_MINUTE = 0;
    private static final int TOP_OF_HOUR = 1;
    private static final int HALF_DAY = 2;
    private static final int TOP_OF_DAY = 3;
    private static final int TOP_OF_WEEK = 4;
    private static final int TOP_OF_MONTH = 5;

    /**
     * The date pattern. By default, the pattern is set to "'.'yyyy-MM-dd" meaning daily rollover.
     */
    private String datePattern = "'.'yyyy-MM-dd";
    private boolean compressBackups = false;
    private int maxDays = 7;

    /**
     * The log file will be renamed to the value of the scheduledFilename variable when the next interval is entered.
     * For example, if the rollover period is one hour, the log file will be renamed to the value of "scheduledFilename"
     * at the beginning of the next hour.
     *
     * The precise time when a rollover occurs depends on logging activity.
     */
    private String scheduledFilename;

    /**
     * The next time we estimate a rollover should occur.
     */
    private long nextCheck = System.currentTimeMillis() - 1;
    Date lastCheck = new Date();
    SimpleDateFormat datePatternFormat = new SimpleDateFormat(datePattern);
    RollingCalendar rollingCalendar = new RollingCalendar();
    int checkPeriod = TOP_OF_TROUBLE;

    /**
     * The default constructor does nothing.
     */
    public CustodianDailyRollingFileAppender() {
    }

    /**
     * Instantiate a <code>DailyRollingFileAppender</code> and open the file designated by <code>filename</code>. The
     * opened filename will become the ouptut destination for this appender.
     *
     */
    public CustodianDailyRollingFileAppender(Layout layout, String filename,
        String datePattern) throws IOException {
        super(layout, filename, true);
        this.datePattern = datePattern;
        activateOptions();
    }

    /**
     * Make sure the directory structure for the file exists
     * Based on JBoss DailyRollingFileAppender, org.jboss.logging.appender.DailyRollingFileAppender
     *
     */
    public void setFile(final String filename) {
        makePath(filename);
        super.setFile(filename);
    }

    public void makePath(final String filename) {
        File dir;

        try {
            URL url = new URL(filename.trim());
            dir = new File(url.getFile()).getParentFile();
        } catch (MalformedURLException e) {
            dir = new File(filename.trim()).getParentFile();
        }

        if (!dir.exists()) {
            boolean success = dir.mkdirs();

            if (!success) {
                LogLog.error("Failed to create directory structure: " + dir);
            }
        }
    }

    /**
     * The <b>DatePattern</b> takes a string in the same format as expected by {@link SimpleDateFormat}. This options
     * determines the rollover schedule.
     */
    public void setDatePattern(String pattern) {
        datePattern = pattern;
    }

    /** Returns the value of the <b>DatePattern</b> option. */
    public String getDatePattern() {
        return datePattern;
    }

    public void activateOptions() {
        super.activateOptions();

        if ((datePattern != null) && (fileName != null)) {
            lastCheck.setTime(System.currentTimeMillis());
            datePatternFormat = new SimpleDateFormat(datePattern);

            int type = computeCheckPeriod();
            printPeriodicity(type);
            rollingCalendar.setType(type);

            File file = new File(fileName);
            scheduledFilename = fileName +
                datePatternFormat.format(new Date(file.lastModified()));
        } else {
            LogLog.error(
                "Either File or DatePattern options are not set for appender [" +
                name + "].");
        }
    }

    void printPeriodicity(int type) {
        switch (type) {
        case TOP_OF_MINUTE:
            LogLog.debug("Appender [" + name + "] to be rolled every minute.");

            break;

        case TOP_OF_HOUR:
            LogLog.debug("Appender [" + name +
                "] to be rolled on top of every hour.");

            break;

        case HALF_DAY:
            LogLog.debug("Appender [" + name +
                "] to be rolled at midday and midnight.");

            break;

        case TOP_OF_DAY:
            LogLog.debug("Appender [" + name + "] to be rolled at midnight.");

            break;

        case TOP_OF_WEEK:
            LogLog.debug("Appender [" + name +
                "] to be rolled at start of week.");

            break;

        case TOP_OF_MONTH:
            LogLog.debug("Appender [" + name +
                "] to be rolled at start of every month.");

            break;

        default:
            LogLog.warn("Unknown periodicity for appender [" + name + "].");
        }
    }

    // This method computes the roll over period by looping over the
    // periods, starting with the shortest, and stopping when the r0 is
    // different from from r1, where r0 is the epoch formatted according
    // the datePattern (supplied by the user) and r1 is the
    // epoch+nextMillis(i) formatted according to datePattern. All date
    // formatting is done in GMT and not local format because the test
    // logic is based on comparisons relative to 1970-01-01 00:00:00
    // GMT (the epoch).
    int computeCheckPeriod() {
        RollingCalendar rollingCalendar = new RollingCalendar(GMT,
                Locale.ENGLISH);

        // set sate to 1970-01-01 00:00:00 GMT
        Date epoch = new Date(0);

        if (datePattern != null) {
            for (int i = TOP_OF_MINUTE; i <= TOP_OF_MONTH; i++) {
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(datePattern);
                simpleDateFormat.setTimeZone(GMT); // do all date
                                                   // formatting in GMT

                String r0 = simpleDateFormat.format(epoch);
                rollingCalendar.setType(i);

                Date next = new Date(rollingCalendar.getNextCheckMillis(epoch));
                String r1 = simpleDateFormat.format(next);

                if ((r0 != null) && (r1 != null) && !r0.equals(r1)) {
                    return i;
                }
            }
        }

        return TOP_OF_TROUBLE; // Deliberately head for trouble...
    }

    /**
     * Rollover the current file to a new file.
     */
    void rollOver() throws IOException {
        /* Compute filename, but only if datePattern is specified */
        if (datePattern == null) {
            errorHandler.error("Missing DatePattern option in rollOver().");

            return;
        }

        String datedFilename = fileName + datePatternFormat.format(lastCheck);

        // It is too early to roll over because we are still within the
        // bounds of the current interval. Rollover will occur once the
        // next interval is reached.
        if (scheduledFilename.equals(datedFilename)) {
            return;
        }

        // close current file, and rename it to datedFilename
        this.closeFile();

        File target = new File(scheduledFilename);

        if (target.exists()) {
            target.delete();
        }

        File file = new File(fileName);
        boolean result = file.renameTo(target);

        if (result) {
            LogLog.debug(fileName + " -> " + scheduledFilename);
        } else {
            LogLog.error("Failed to rename [" + fileName + "] to [" +
                scheduledFilename + "].");
        }

        try {
            // This will also close the file. This is OK since multiple
            // close operations are safe.
            this.setFile(fileName, false, this.bufferedIO, this.bufferSize);
        } catch (IOException e) {
            errorHandler.error("setFile(" + fileName + ", false) call failed.");
        }

        scheduledFilename = datedFilename;
    }

    /**
     * This method differentiates CustodianDailyRollingFileAppender from its super class.
     *
     * <p>
     * Before actually logging, this method will check whether it is time to do a rollover. If it is, it will schedule
     * the next rollover time and then rollover.
     */
    protected void subAppend(LoggingEvent event) {
        long n = System.currentTimeMillis();

        if (n >= nextCheck) {
            lastCheck.setTime(n);
            nextCheck = rollingCalendar.getNextCheckMillis(lastCheck);

            try {
                cleanupAndRollOver();
            } catch (IOException ioe) {
                LogLog.error("rollOver() failed.", ioe);
            }
        }

        super.subAppend(event);
    }

    /**
     * This method checks to see if we're exceeding the number of log backups that we are supposed to keep, and if so,
     * deletes the offending files. It then delegates to the rollover method to rollover to a new file if required.
     */
    protected void cleanupAndRollOver() throws IOException {
        // Check to see if there are already 5 files
        File file = new File(fileName);
        Calendar cal = Calendar.getInstance();
        cal.add(Calendar.DATE, -getMaxDays());

        Date cutoffDate = cal.getTime();

        if (file.getParentFile().exists()) {
            File[] files = file.getParentFile()
                               .listFiles(new StartsWithFileFilter(
                        file.getName(), false));
            int nameLength = file.getName().length();

            for (int i = 0; i < files.length; i++) {
                String datePart = null;

                try {
                    // The datePart is appended to the end, so everything that
                    // is not the fixed name part is the DF.
                    datePart = files[i].getName().substring(nameLength);

                    Date date = datePatternFormat.parse(datePart);

                    if (date.before(cutoffDate)) {
                        files[i].delete();
                    }
                    // If we're supposed to zip files and this isn't already a
                    // zip
                    else if (shouldCompress()) {
                        zipAndDelete(files[i]);
                    }
                } catch (Exception pe) {
                    // This isn't a file we should touch (it isn't named
                    // correctly)
                }
            }
        }

        rollOver();
    }

    /**
     * Compresses the passed file to a .zip file, stores the .zip in the same directory as the passed file, and then
     * deletes the original, leaving only the .zipped archive.
     *
     * @param file
     */
    private void zipAndDelete(File file) throws IOException {
        if (!file.getName().endsWith(".zip")) {
            File zipFile = new File(file.getParent(), file.getName() + ".zip");
            FileInputStream fis = new FileInputStream(file);
            FileOutputStream fos = new FileOutputStream(zipFile);
            ZipOutputStream zos = new ZipOutputStream(fos);
            ZipEntry zipEntry = new ZipEntry(file.getName());
            zos.putNextEntry(zipEntry);

            byte[] buffer = new byte[4096];

            while (true) {
                int bytesRead = fis.read(buffer);

                if (bytesRead == -1) {
                    break;
                } else {
                    zos.write(buffer, 0, bytesRead);
                }
            }

            zos.closeEntry();
            fis.close();
            zos.close();
            file.delete();
        }
    }

    public String getMaxNumberOfDays() {
        return String.valueOf(maxDays);
    }

    public int getMaxDays() {
        return maxDays;
    }

    public void setMaxNumberOfDays(String maxNumberOfDays) {
        try {
            this.maxDays = Integer.parseInt(maxNumberOfDays);
        } catch (NumberFormatException e) {
            LogLog.warn("Could not parse numberOfDays-string " +
                maxNumberOfDays + ", using 7 as default", e);
            this.maxDays = 7;
        }
    }

    public boolean shouldCompress() {
        return compressBackups;
    }

    // Needed to be able to set this property..log4j probably uses BeanUtils, which requires this :-)
    public String getCompressBackups() {
        return String.valueOf(compressBackups);
    }

    public void setCompressBackups(String compressBackups) {
        this.compressBackups = "YES".equalsIgnoreCase(compressBackups) ||
            "TRUE".equalsIgnoreCase(compressBackups);
    }

    class StartsWithFileFilter implements FileFilter {
        private String startsWith;
        private boolean inclDirs = false;

        /**
         *
         */
        public StartsWithFileFilter(String startsWith,
            boolean includeDirectories) {
            super();
            this.startsWith = startsWith.toUpperCase();
            inclDirs = includeDirectories;
        }

        /*
         * (non-Javadoc)
         *
         * @see java.io.FileFilter#accept(java.io.File)
         */
        public boolean accept(File pathname) {
            if (!inclDirs && pathname.isDirectory()) {
                return false;
            } else {
                return pathname.getName().toUpperCase().startsWith(startsWith);
            }
        }
    }

    /**
     * RollingCalendar is a helper class to CustodianDailyRollingFileAppender. Given a periodicity type and the current
     * time, it computes the start of the next interval.
     */
    class RollingCalendar extends GregorianCalendar {
        private static final long serialVersionUID = -3560331770601814177L;
        int type = CustodianDailyRollingFileAppender.TOP_OF_TROUBLE;

        RollingCalendar() {
            super();
        }

        RollingCalendar(TimeZone tz, Locale locale) {
            super(tz, locale);
        }

        void setType(int type) {
            this.type = type;
        }

        public long getNextCheckMillis(Date now) {
            return getNextCheckDate(now).getTime();
        }

        public Date getNextCheckDate(Date now) {
            this.setTime(now);

            switch (type) {
            case CustodianDailyRollingFileAppender.TOP_OF_MINUTE:
                this.set(Calendar.SECOND, 0);
                this.set(Calendar.MILLISECOND, 0);
                this.add(Calendar.MINUTE, 1);

                break;

            case CustodianDailyRollingFileAppender.TOP_OF_HOUR:
                this.set(Calendar.MINUTE, 0);
                this.set(Calendar.SECOND, 0);
                this.set(Calendar.MILLISECOND, 0);
                this.add(Calendar.HOUR_OF_DAY, 1);

                break;

            case CustodianDailyRollingFileAppender.HALF_DAY:
                this.set(Calendar.MINUTE, 0);
                this.set(Calendar.SECOND, 0);
                this.set(Calendar.MILLISECOND, 0);

                int hour = get(Calendar.HOUR_OF_DAY);

                if (hour < 12) {
                    this.set(Calendar.HOUR_OF_DAY, 12);
                } else {
                    this.set(Calendar.HOUR_OF_DAY, 0);
                    this.add(Calendar.DAY_OF_MONTH, 1);
                }

                break;

            case CustodianDailyRollingFileAppender.TOP_OF_DAY:
                this.set(Calendar.HOUR_OF_DAY, 0);
                this.set(Calendar.MINUTE, 0);
                this.set(Calendar.SECOND, 0);
                this.set(Calendar.MILLISECOND, 0);
                this.add(Calendar.DATE, 1);

                break;

            case CustodianDailyRollingFileAppender.TOP_OF_WEEK:
                this.set(Calendar.DAY_OF_WEEK, getFirstDayOfWeek());
                this.set(Calendar.HOUR_OF_DAY, 0);
                this.set(Calendar.MINUTE, 0);
                this.set(Calendar.SECOND, 0);
                this.set(Calendar.MILLISECOND, 0);
                this.add(Calendar.WEEK_OF_YEAR, 1);

                break;

            case CustodianDailyRollingFileAppender.TOP_OF_MONTH:
                this.set(Calendar.DATE, 1);
                this.set(Calendar.HOUR_OF_DAY, 0);
                this.set(Calendar.MINUTE, 0);
                this.set(Calendar.SECOND, 0);
                this.set(Calendar.MILLISECOND, 0);
                this.add(Calendar.MONTH, 1);

                break;

            default:
                throw new IllegalStateException("Unknown periodicity type.");
            }

            return getTime();
        }
    }
}
