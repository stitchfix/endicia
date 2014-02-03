require 'rubygems'
require 'httparty'
require 'active_support/core_ext'
require 'builder'
require 'uri'

require 'endicia/label'
require 'endicia/rails_helper'

# Hack fix because Endicia sends response back without protocol in xmlns uri
module HTTParty
  class Request
    alias_method :parse_response_without_hack, :parse_response
    def parse_response(body)
      Rails.logger.info("RESPONSE>")
      Rails.logger.info(body.sub(/<Base64LabelImage>.*<\/Base64LabelImage>/, '<Base64LabelImage>(image data stripped)</Base64LabelImage>'))
      Rails.logger.info("<RESPONSE")

      parse_response_without_hack(
        body.sub(/xmlns=("|')(www.envmgr.com|LabelServer.Endicia.com)/, 'xmlns=\1https://\2'))
    end
  end
end

module Endicia
  include HTTParty
  extend RailsHelper

  class EndiciaError < StandardError; end
  class InsuranceError < EndiciaError; end

  JEWELRY_INSURANCE_EXCLUDED_ZIPS = %w(10036 10017 94102 94108)

  # We need the following to make requests
  # RequesterID (string): Requester ID (also called Partner ID) uniquely identifies the system making the request. Endicia assigns this ID. The Test Server does not authenticate the RequesterID. Any text value of 1 to 50 characters is valid.
  # AccountID (6 digits): Account ID for the Endicia postage account. The Test Server does not authenticate the AccountID. Any 6-digit value is valid.
  # PassPhrase (string): Pass Phrase for the Endicia postage account. The Test Server does not authenticate the PassPhrase. Any text value of 1 to 64 characters is valid.

  # We probably want the following arguments
  # MailClass, WeightOz, MailpieceShape, Machinable, FromPostalCode

  format :xml
  # example XML
  # <LabelRequest><ReturnAddress1>884 Railroad Street, Suite C</ReturnAddress1><ReturnCity>Ypsilanti</ReturnCity><ReturnState>MI</ReturnState><FromPostalCode>48197</FromPostalCode><FromCity>Ypsilanti</FromCity><FromState>MI</FromState><FromCompany>VGKids</FromCompany><ToPostalCode>48197</ToPostalCode><ToAddress1>1237 Elbridge St</ToAddress1><ToCity>Ypsilanti</ToCity><ToState>MI</ToState><PartnerTransactionID>123</PartnerTransactionID><PartnerCustomerID>71212</PartnerCustomerID><MailClass>MediaMail</MailClass><Test>YES</Test><RequesterID>poopants</RequesterID><AccountID>792190</AccountID><PassPhrase>whiplash1</PassPhrase><WeightOz>10</WeightOz></LabelRequest>

  # Request a shipping label.
  #
  # Accepts a hash of options in the form:
  # { :NodeOrAttributeName => "value", ... }
  #
  # See https://app.sgizmo.com/users/4508/Endicia_Label_Server.pdf Table 3-1
  # for available options.
  #
  # Note: options should be specified in a "flat" hash, they should not be
  # formated to fit the nesting of the XML.
  #
  # If you are using rails, any applicable options specified in
  # config/endicia.yml will be used as defaults. For example:
  #
  #     development:
  #       Test: YES
  #       AccountID: 123
  #       ...
  #
  # Returns a Endicia::Label object.
  def self.get_label(opts={})
    opts = defaults.merge(opts)
    opts[:Test] ||= "NO"
    url = "#{label_service_url(opts)}/GetPostageLabelXML"
    opts = clean_options(opts)
    insurance = extract_insurance(opts)
    handle_extended_zip_code(opts)

    root_keys = :LabelType, :Test, :LabelSize, :ImageFormat, :ImageResolution
    root_attributes = extract(opts, root_keys)
    root_attributes[:LabelType] ||= "Default"

    dimension_keys = :Length, :Width, :Height
    mailpiece_dimenions = extract(opts, dimension_keys)

    xml = Builder::XmlMarkup.new
    body = "labelRequestXML=" + xml.LabelRequest(root_attributes) do |xm|
      opts.each { |key, value| xm.tag!(key, value) }
      xm.Services({ :InsuredMail => insurance }) if insurance
      unless mailpiece_dimenions.empty?
        xm.MailpieceDimensions do |md|
          mailpiece_dimenions.each { |key, value| md.tag!(key, value) }
        end
      end
    end

    result = self.post(url, :body => body)
    Endicia::Label.new(result).tap do |the_label|
      the_label.request_body = body.to_s
      the_label.request_url = url
    end
  end

  # Change your account pass phrase. This is a required step to move to
  # production use after requesting an account.
  #
  # Accepts the new phrase and a hash of options in the form:
  #
  #     { :Name => "value", ... }
  #
  # See https://app.sgizmo.com/users/4508/Endicia_Label_Server.pdf Table 5-1
  # for available/required options.
  #
  # Note: options should be specified in a "flat" hash, they should not be
  # formated to fit the nesting of the XML.
  #
  # If you are using rails, any applicable options specified in
  # config/endicia.yml will be used as defaults. For example:
  #
  #     development:
  #       Test: YES
  #       AccountID: 123
  #       ...
  #
  # Returns a hash in the form:
  #
  #     {
  #       :success => true, # or false
  #       :error_message => "the message", # or nil
  #       :response_body => "the response body"
  #     }
  def self.change_pass_phrase(new_phrase, options = {})
    url = "#{label_service_url(options)}/ChangePassPhraseXML"
    options = clean_options(options)

    xml = Builder::XmlMarkup.new
    body = "changePassPhraseRequestXML=" + xml.ChangePassPhraseRequest do |xml|
      authorize_request(xml, options)
      xml.NewPassPhrase new_phrase
      xml.RequestID "CPP#{Time.now.to_f}"
    end

    result = self.post(url, { :body => body })
    parse_result(result, "ChangePassPhraseRequestResponse")
  end

  # Add postage to your account (submit a RecreditRequest). This is a required
  # step to move to production use after requesting an account and changing
  # your pass phrase.
  #
  # Accepts the amount (in dollars) and a hash of options in the form:
  #
  #     { :Name => "value", ... }
  #
  # See https://app.sgizmo.com/users/4508/Endicia_Label_Server.pdf Table 5-1
  # for available/required options.
  #
  # Note: options should be specified in a "flat" hash, they should not be
  # formated to fit the nesting of the XML.
  #
  # If you are using rails, any applicable options specified in
  # config/endicia.yml will be used as defaults. For example:
  #
  #     development:
  #       Test: YES
  #       AccountID: 123
  #       ...
  #
  # Returns a hash in the form:
  #
  #     {
  #       :success => true, # or false
  #       :error_message => "the message", # or nil if no error
  #       :response_body => "the response body"
  #     }
  def self.buy_postage(amount, options = {})

    url = "#{label_service_url(options)}/BuyPostageXML"
    options = clean_options(options)

    xml = Builder::XmlMarkup.new
    body = "recreditRequestXML=" + xml.RecreditRequest do |xml|
      authorize_request(xml, options)
      xml.RecreditAmount amount
      xml.RequestID "BP#{Time.now.to_f}"
    end
    result = self.post(url, { :body => body })
    parse_result(result, "RecreditRequestResponse")
  end

  # Given a tracking number, return a status message for the shipment.
  #
  # See https://app.sgizmo.com/users/4508/Endicia_Label_Server.pdf Table 12-1
  # for available/required options.
  #
  # Note: options should be specified in a "flat" hash, they should not be
  # formated to fit the nesting of the XML.
  #
  # If you are using rails, any applicable options specified in
  # config/endicia.yml will be used as defaults. For example:
  #
  #     development:
  #       Test: YES
  #       AccountID: 123
  #       ...
  #
  # Returns a hash in the form:
  #
  #     {
  #       :success => true, # or false
  #       :error_message => "the message", # or nil if no error
  #       :status => "the package status", # or nil if error
  #       :response_body => "the response body"
  #     }
  def self.status_request(tracking_number, options = {})
    options = clean_options(options)
    xml = Builder::XmlMarkup.new.StatusRequest do |xml|
      xml.AccountID(options[:AccountID] || defaults[:AccountID])
      xml.PassPhrase(options[:PassPhrase] || defaults[:PassPhrase])
      xml.Test(options[:Test] || defaults[:Test] || "NO")
      xml.FullStatus(options[:FullStatus] || defaults[:FullStatus] || '')
      xml.StatusList { |xml| xml.PICNumber(tracking_number) }
    end

    if options[:logger]
      options[:logger].info("ENDICIA REQUEST: #{tracking_number}")
      options[:logger].info("\n[REQUEST]")
      options[:logger].info(xml)
      options[:logger].info("[ENDREQUEST]")
    end

    params = { :method => 'StatusRequest', :XMLInput => URI.encode(xml) }
    result = self.get(els_service_url(params))
    response_body = result.body
    response_body.gsub!(/<PICNumber>[^<]*/, "<PICNumber>")
    response = {
      :success => false,
      :error_message => nil,
      :status => nil,
      :response_body => response_body
    }

    if options[:logger]
      options[:logger].info("\n[RESPONSE]")
      options[:logger].info(xml)
      options[:logger].info("[ENDRESPONSE]")
    end

    # TODO: It is possible to make a batch status request, currently this only
    #       supports one at a time. The response that comes back is not parsed
    #       well by HTTParty. So we have to assume there is only one tracking
    #       number in order to parse it with a regex.

    if result && result = result['StatusResponse']
      unless response[:error_message] = result['ErrorMsg']
        response[:status] = response_body.match(/<Status>(.+)<\/Status>/)[1]
        status_code = response_body.match(/<StatusCode>(.+)<\/StatusCode>/)[1]
        response[:success] = (status_code.to_s != '-1')
      end
    end

    response
  end

  # Given a tracking number, try and void the label generated in a previous call
  #
  # See https://app.sgizmo.com/users/4508/Endicia_Label_Server.pdf Table 11-1
  # for available/required options.
  #
  # Note: options should be specified in a "flat" hash, they should not be
  # formated to fit the nesting of the XML.
  #
  # If you are using rails, any applicable options specified in
  # config/endicia.yml will be used as defaults. For example:
  #
  #     development:
  #       Test: YES
  #       AccountID: 123
  #       ...
  #
  # Returns a hash in the form:
  #
  #     {
  #       :success          => true,            # if all tracking numbers are approved or false
  #                                             # Probably better to use tracking_number array
  #       :error_message    => "the message",   # DEPRECATED message describing success or failure.
  #                                             # Should instead look for messages in the :tracking_numbers array
  #       :form_number      => 12345,           # Form Number for refunded label
  #       :tracking_numbers => [                # An array with information about all requested tracking numbers
  #         :pic_number => '123456789',         # the tracking number you requested
  #         :approved => true,                  # or false
  #         :message => "the message"           # message describing success or failure
  #       ]
  #       :response_body => "the response body" # the raw HTTP response
  #     }
  def self.refund_request(tracking_number, options = {})
    options = clean_options(options)
    # If we didn't get an array of tracking numbers make it one for simplicity
    tracking_numbers = tracking_number.is_a?(Array) ? tracking_number : [tracking_number]

    xml = Builder::XmlMarkup.new.RefundRequest do |xml|
      xml.AccountID(options[:AccountID] || defaults[:AccountID])
      xml.PassPhrase(options[:PassPhrase] || defaults[:PassPhrase])
      xml.Test(options[:Test] || defaults[:Test] || "NO")
      xml.RefundList { |xml| tracking_numbers.collect{ |tracking_number| xml.PICNumber(tracking_number) } }
    end

    params = { :method => 'RefundRequest', :XMLInput => URI.encode(xml) }
    result = self.get(els_service_url(params))

    response = {
      :success => true,
      :error_message => nil,
      :response_body => result.body,
      :tracking_numbers => []
    }

    if result && result = result['RefundResponse']
      unless response[:error_message] = result['ErrorMsg']
        response[:form_number]   = result['FormNumber']

        result = result['RefundList']['PICNumber']
        if result
          # If we didn't get an array make it one for simplicity
          result = result.is_a?(Array) ? result : [result]
          result.each do |r|
            approved = r['IsApproved']['__content__'] == 'YES' ? true : false
            response[:tracking_numbers] << {
              :pic_number => r['__content__'].try(:strip),
              :approved => approved,
              :message => r['ErrorMsg']['__content__'].try(:strip)
            }
            response[:success] = response[:success] && approved
          end
        end
        # set error_message to the first one in tracking numbers array
        response[:error_message] = response[:tracking_numbers].first[:message] rescue nil
      end
    end

    response
  end


  # Given a tracking number and package location code,
  # return a carrier pickup confirmation.
  #
  # See https://app.sgizmo.com/users/4508/Endicia_Label_Server.pdf Table 15-1
  # for available/required options, and package location codes.
  #
  # If you are using rails, any applicable options specified in
  # config/endicia.yml will be used as defaults. For example:
  #
  #     development:
  #       Test: YES
  #       AccountID: 123
  #       ...
  #
  # Returns a hash in the form:
  #
  #     {
  #       :success => true, # or false
  #       :error_message => "the message", # or nil if no error message
  #       :error_code => "usps error code", # or nil if no error
  #       :error_description => "usps error description", # or nil if no error
  #       :day_of_week => "pickup day of week (ex: Monday)",
  #       :date => "xx/xx/xxxx", # date of pickup,
  #       :confirmation_number => "confirmation number of the pickup", # save this!
  #       :response_body => "the response body"
  #     }
  def self.carrier_pickup_request(tracking_number, package_location, options = {})
    options = clean_options(options)
    xml = Builder::XmlMarkup.new.CarrierPickupRequest do |xml|
      xml.AccountID(options.delete(:AccountID) || defaults[:AccountID])
      xml.PassPhrase(options.delete(:PassPhrase) || defaults[:PassPhrase])
      xml.Test(options.delete(:Test) || defaults[:Test] || "NO")
      xml.PackageLocation(package_location)
      xml.PickupList { |xml| xml.PICNumber(tracking_number) }
      options.each { |key, value| xml.tag!(key, value) }
    end

    params = { :method => 'CarrierPickupRequest', :XMLInput => URI.encode(xml) }
    result = self.get(els_service_url(params))

    response = {
      :success => false,
      :response_body => result.body
    }

    # TODO: this is some nasty logic...
    if result && result = result["CarrierPickupRequestResponse"]
      unless response[:error_message] = result['ErrorMsg']
        if result = result["Response"]
          if error = result.delete("Error")
            response[:error_code] = error["Number"]
            response[:error_description] = error["Description"]
          else
            response[:success] = true
          end
          result.each { |key, value| response[key.underscore.to_sym] = value }
        end
      end
    end

    response
  end

  private

  def self.extract(hash, keys)
    {}.tap do |return_hash|
      keys.each do |key|
        value = return_hash[key] = hash.delete(key)
        return_hash.delete(key) if value.nil? || value.empty?
      end
    end
  end

  # Given a builder object, add the auth nodes required for many api calls.
  # Will pull values from options hash or defaults if not found.
  def self.authorize_request(xml_builder, options = {})
    requester_id = options[:RequesterID] || defaults[:RequesterID]
    account_id   = options[:AccountID]   || defaults[:AccountID]
    pass_phrase  = options[:PassPhrase]  || defaults[:PassPhrase]

    xml_builder.RequesterID requester_id
    xml_builder.CertifiedIntermediary do |xml_builder|
      xml_builder.AccountID account_id
      xml_builder.PassPhrase pass_phrase
    end
  end

  # Return the url for making requests.
  # Pass options hash with :Test => "YES" to return the url of the test server
  # (this matches the Test attribute/node value for most API calls).
  def self.label_service_url(options = {})
    if options[:Test] == "SANDBOX"
      url = "https://elstestserver.endicia.com"
    else
      test = (options[:Test] || defaults[:Test] || "NO").upcase == "YES"
      url = test ? "https://www.envmgr.com" : "https://LabelServer.Endicia.com"
    end
    "#{url}/LabelService/EwsLabelService.asmx"
  end

  # Some requests use the ELS service url. This URL is used for requests that
  # can accept GET, and have params passed via URL instead of a POST body.
  # Pass a hash of params to have them converted to a &key=value string and
  # appended to the URL.
  def self.els_service_url(params = {})
    params = params.to_a.map { |i| "#{i[0]}=#{i[1]}"}.join('&')
    "http://www.endicia.com/ELS/ELSServices.cfc?wsdl&#{params}"
  end

  def self.defaults
    if rails? && @defaults.nil?
      config_file = File.join(rails_root, 'config', 'endicia.yml')
      if File.exist?(config_file)
        @defaults = YAML.load_file(config_file)[rails_env].symbolize_keys
      end
    end

    @defaults || {}
  end

  def self.clean_options(options)
    options[:Test] = 'YES' if options[:Test] == 'SANDBOX'
    options
  end

  def self.parse_result(result, root)
    parsed_result = {
      :success => false,
      :error_message => nil,
      :response_body => result.body
    }

    if result && result[root]
      root = result[root]
      parsed_result[:error_message] = root["ErrorMessage"]
      parsed_result[:success] = root["Status"] && root["Status"].to_s == "0"
    end

    parsed_result
  end

  # Handle special case where jewelry can't have insurance if sent to certain zips
  def self.extract_insurance(opts)
    jewelry = opts.delete(:Jewelry)
    opts.delete(:InsuredMail).tap do |insurance|
      if insurance && insurance == "Endicia" && jewelry
        if JEWELRY_INSURANCE_EXCLUDED_ZIPS.include? opts[:ToPostalCode]
          raise InsuranceError, "Can't ship jewelry with insurance to #{opts[:ToPostalCode]}"
        end
      end
    end
  end

  def self.handle_extended_zip_code(opts)
    if m = /([0-9]{5})-([0-9]{4})/.match(opts[:ToPostalCode])
      opts[:ToPostalCode] = m[1]
      opts[:ToZIP4] = m[2]
    end
  end
end
